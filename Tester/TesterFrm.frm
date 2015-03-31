VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Tester 
   Caption         =   "ALCOR TESTER"
   ClientHeight    =   10290
   ClientLeft      =   2250
   ClientTop       =   645
   ClientWidth     =   13185
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "TesterFrm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10290
   ScaleWidth      =   13185
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   2280
      TabIndex        =   83
      Top             =   840
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   2280
      Top             =   1440
   End
   Begin VB.TextBox SkipOtherDevice 
      Height          =   375
      Left            =   2400
      TabIndex        =   81
      Text            =   "enter skip number"
      Top             =   2280
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.ComboBox SkipCtrlCount 
      Height          =   360
      ItemData        =   "TesterFrm.frx":0CCA
      Left            =   6960
      List            =   "TesterFrm.frx":0CEC
      Style           =   2  '單純下拉式
      TabIndex        =   79
      Top             =   600
      Width           =   615
   End
   Begin VB.CheckBox SaveOSLog 
      Caption         =   "Save O/S Log"
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
      Left            =   2400
      TabIndex        =   76
      Top             =   2880
      Value           =   1  '核取
      Width           =   1815
   End
   Begin VB.CheckBox OSCheck 
      Caption         =   "OS_Rec"
      Height          =   615
      Left            =   11880
      TabIndex        =   75
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CheckBox AudioRecordCheck 
      Caption         =   "Check1"
      Height          =   255
      Left            =   11760
      TabIndex        =   74
      Top             =   7440
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   4800
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   73
      Top             =   1440
      Width           =   375
   End
   Begin MCI.MMControl mmcAudio 
      Height          =   495
      Left            =   7680
      TabIndex        =   72
      Top             =   1560
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   873
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      Height          =   735
      Left            =   11400
      TabIndex        =   71
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox Text25 
      Height          =   360
      Left            =   6600
      TabIndex        =   69
      Text            =   "3"
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton Rec_Mp34 
      Caption         =   "Rec Mp34"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11880
      TabIndex        =   67
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton Rec_Mp33 
      Caption         =   "Rec Mp33"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11880
      TabIndex        =   66
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Rec_MP32 
      Caption         =   "Rec Mp32"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11880
      TabIndex        =   65
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Rec_MP31 
      Caption         =   "Rec Mp31"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11880
      TabIndex        =   64
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton MP3_Rec 
      Caption         =   "Rec MP3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11880
      TabIndex        =   63
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "ShowTable"
      Height          =   735
      Left            =   11520
      TabIndex        =   60
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox Text24 
      Height          =   375
      Left            =   9960
      TabIndex        =   58
      Text            =   "Text24"
      Top             =   8880
      Width           =   1335
   End
   Begin VB.TextBox Text23 
      Height          =   375
      Left            =   9960
      TabIndex        =   56
      Text            =   "Text23"
      Top             =   7920
      Width           =   1335
   End
   Begin VB.TextBox Text22 
      Height          =   375
      Left            =   9960
      TabIndex        =   54
      Text            =   "Text22"
      Top             =   6960
      Width           =   1335
   End
   Begin VB.TextBox Text21 
      Height          =   375
      Left            =   9960
      TabIndex        =   52
      Text            =   "Text21"
      Top             =   6000
      Width           =   1335
   End
   Begin VB.TextBox Text20 
      Height          =   375
      Left            =   8400
      TabIndex        =   50
      Text            =   "Text20"
      Top             =   8880
      Width           =   1335
   End
   Begin VB.TextBox Text19 
      Height          =   375
      Left            =   8400
      TabIndex        =   48
      Text            =   "Text19"
      Top             =   7920
      Width           =   1335
   End
   Begin VB.TextBox Text18 
      Height          =   360
      Left            =   8400
      TabIndex        =   46
      Text            =   "Text18"
      Top             =   6960
      Width           =   1335
   End
   Begin VB.TextBox Text17 
      Height          =   375
      Left            =   8400
      TabIndex        =   44
      Text            =   "Text17"
      Top             =   6000
      Width           =   1335
   End
   Begin VB.TextBox Text16 
      Height          =   375
      Left            =   6960
      TabIndex        =   42
      Text            =   "Text16"
      Top             =   8880
      Width           =   1215
   End
   Begin VB.TextBox Text15 
      Height          =   375
      Left            =   6960
      TabIndex        =   40
      Text            =   "Text15"
      Top             =   7920
      Width           =   1215
   End
   Begin VB.TextBox Text14 
      Height          =   375
      Left            =   6960
      TabIndex        =   39
      Text            =   "Text14"
      Top             =   6960
      Width           =   1215
   End
   Begin VB.TextBox Text13 
      Height          =   375
      Left            =   6960
      TabIndex        =   38
      Text            =   "Text13"
      Top             =   6000
      Width           =   1215
   End
   Begin VB.TextBox Text12 
      Height          =   375
      Left            =   5640
      TabIndex        =   34
      Text            =   "Text12"
      Top             =   8880
      Width           =   1215
   End
   Begin VB.TextBox Text11 
      Height          =   375
      Left            =   5640
      TabIndex        =   33
      Text            =   "Text11"
      Top             =   7920
      Width           =   1215
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   5640
      TabIndex        =   32
      Text            =   "Text10"
      Top             =   6960
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   5640
      TabIndex        =   31
      Text            =   "Text9"
      Top             =   6000
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   9960
      TabIndex        =   26
      Text            =   "Text8"
      Top             =   4920
      Width           =   1335
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   8520
      TabIndex        =   25
      Text            =   "Text7"
      Top             =   4920
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   7080
      TabIndex        =   24
      Text            =   "Text6"
      Top             =   4920
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   3960
      TabIndex        =   23
      Text            =   "Text4"
      Top             =   4920
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   3960
      TabIndex        =   18
      Text            =   "Text2"
      Top             =   5520
      Width           =   1455
   End
   Begin VB.TextBox txtmsg 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   2400
      MultiLine       =   -1  'True
      TabIndex        =   17
      Text            =   "TesterFrm.frx":0D0E
      Top             =   6240
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   2400
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   5520
      Width           =   1455
   End
   Begin MSCommLib.MSComm MSComm2 
      Left            =   10320
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   9720
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
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
      Left            =   5640
      TabIndex        =   12
      Text            =   "Text5"
      Top             =   4920
      Width           =   1335
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
      Height          =   375
      Left            =   2400
      TabIndex        =   11
      Text            =   "Text3"
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "關閉程式"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9240
      TabIndex        =   10
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   10920
      Top             =   1080
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "啟動測試端功能"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6960
      MaskColor       =   &H0080C0FF&
      TabIndex        =   0
      Top             =   2040
      Width           =   2295
   End
   Begin VB.CheckBox FixCardMode 
      Caption         =   "Fix Card mode"
      Height          =   255
      Left            =   7800
      TabIndex        =   70
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label38 
      Caption         =   "3510_Test_Progress"
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
      Left            =   2280
      TabIndex        =   82
      Top             =   600
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Version_Label 
      Alignment       =   2  '置中對齊
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
      Left            =   2520
      TabIndex        =   80
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label Label37 
      Caption         =   "Skip GPIB Counter"
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
      Left            =   4800
      TabIndex        =   78
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label GPIBCARD_Label 
      BackColor       =   &H0080FFFF&
      Caption         =   " GPIB Exist"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   4680
      TabIndex        =   77
      Top             =   0
      Width           =   2775
   End
   Begin VB.Label Label36 
      Caption         =   "Hub Time"
      Height          =   375
      Left            =   5400
      TabIndex        =   68
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label35 
      Caption         =   "Label35"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11520
      TabIndex        =   62
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label34 
      Caption         =   "Label34"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11520
      TabIndex        =   61
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label33 
      BackColor       =   &H00FFFFFF&
      Caption         =   "UNKNOW DEVICE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   2400
      TabIndex        =   59
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Label Label32 
      Caption         =   "Label32"
      Height          =   495
      Left            =   9960
      TabIndex        =   57
      Top             =   8400
      Width           =   1335
   End
   Begin VB.Label Label31 
      Caption         =   "Label31"
      Height          =   495
      Left            =   9960
      TabIndex        =   55
      Top             =   7440
      Width           =   1455
   End
   Begin VB.Label Label30 
      Caption         =   "Label30"
      Height          =   375
      Left            =   9960
      TabIndex        =   53
      Top             =   6480
      Width           =   1335
   End
   Begin VB.Label Label29 
      Caption         =   "Label29"
      Height          =   495
      Left            =   10080
      TabIndex        =   51
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label Label28 
      Caption         =   "Label28"
      Height          =   495
      Left            =   8400
      TabIndex        =   49
      Top             =   8400
      Width           =   1335
   End
   Begin VB.Label Label27 
      Caption         =   "Label27"
      Height          =   495
      Left            =   8400
      TabIndex        =   47
      Top             =   7440
      Width           =   1335
   End
   Begin VB.Label Label26 
      Caption         =   "Label26"
      Height          =   375
      Left            =   8400
      TabIndex        =   45
      Top             =   6480
      Width           =   1335
   End
   Begin VB.Label Label25 
      Caption         =   "Label25"
      Height          =   495
      Left            =   8520
      TabIndex        =   43
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Label Label24 
      Caption         =   "Label24"
      Height          =   495
      Left            =   6960
      TabIndex        =   41
      Top             =   8400
      Width           =   1095
   End
   Begin VB.Label Label23 
      Caption         =   "Label23"
      Height          =   495
      Left            =   6960
      TabIndex        =   37
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Label Label22 
      Caption         =   "Label22"
      Height          =   375
      Left            =   6960
      TabIndex        =   36
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Label Label21 
      Caption         =   "Label21"
      Height          =   495
      Left            =   6960
      TabIndex        =   35
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Label Label20 
      Caption         =   "Label20"
      Height          =   495
      Left            =   5640
      TabIndex        =   30
      Top             =   8400
      Width           =   1215
   End
   Begin VB.Label Label19 
      Caption         =   "Label19"
      Height          =   615
      Left            =   5640
      TabIndex        =   29
      Top             =   7440
      Width           =   1335
   End
   Begin VB.Label Label18 
      Caption         =   "Label18"
      Height          =   375
      Left            =   5640
      TabIndex        =   28
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Label Label17 
      Caption         =   "Label17"
      Height          =   495
      Left            =   5640
      TabIndex        =   27
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Line Line1 
      X1              =   2400
      X2              =   11400
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Label Label16 
      Alignment       =   2  '置中對齊
      Caption         =   "BIN5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10080
      TabIndex        =   22
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label15 
      Alignment       =   2  '置中對齊
      Caption         =   "BIN4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      TabIndex        =   21
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label14 
      Alignment       =   2  '置中對齊
      Caption         =   "BIN3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   20
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label13 
      Alignment       =   2  '置中對齊
      Caption         =   "BIN2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   19
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label12 
      Caption         =   "Label12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4800
      TabIndex        =   15
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label11 
      Alignment       =   2  '置中對齊
      Caption         =   "FAIL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   14
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label Label10 
      Alignment       =   2  '置中對齊
      Caption         =   "PASS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   13
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label Label9 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label9"
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
      Left            =   2400
      TabIndex        =   9
      Top             =   3960
      Width           =   8895
   End
   Begin VB.Label Label8 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00FFFFFF&
      Caption         =   "test result"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   10080
      TabIndex        =   8
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00FFFFFF&
      Caption         =   "step5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   8880
      TabIndex        =   7
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Label6 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00FFFFFF&
      Caption         =   "step4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   7920
      TabIndex        =   6
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00FFFFFF&
      Caption         =   "step2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   5760
      TabIndex        =   5
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00FFFFFF&
      Caption         =   "step3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   6840
      TabIndex        =   4
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00FFFFFF&
      Caption         =   "step1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   3480
      Width           =   975
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
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   2880
      Width           =   6495
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      BackColor       =   &H0000FFFF&
      Caption         =   "PC 測試端 TESTER "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   7680
      TabIndex        =   1
      Top             =   0
      Width           =   3705
   End
End
Attribute VB_Name = "Tester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function waveOutGetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, lpdwVolume As Long) As Long
Private Declare Function waveOutSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, source As Any, ByVal Length As Long)
Const WAVE_MAPPER = -1&

' For startup use
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal Hkey As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal Hkey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Const REG_SZ = 1
Private Const HKEY_LOCAL_MACHINE = &H80000002

Public ContiUnknowFailCounter As Integer
Public ResetHubString As String
Public ResetHubReturn As Long
Public HVFlag As Boolean, LVFlag As Boolean
Public HV_Result As String, LV_Result As String
Public MPFail As Boolean

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
  (ByVal lpClassName As String, _
  ByVal lpWindowName As String) As Long
  
Private Declare Function ShowWindow Lib "user32" _
  (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
  
Private Declare Function SetForegroundWindow Lib "user32" _
  (ByVal hwnd As Long) As Long

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

Public Function WaitProcQuit(pId As Long)

On Error Resume Next
Dim objProcess
Dim Pid_Exist As Boolean

    Do
        Pid_Exist = True
        For Each objProcess In GetObject("winmgmts:\\.\root\cimv2:win32_process").instances_
            'Debug.Print objProcess.Handle; objProcess.Name
             If objProcess.Handle = pId Then
                Pid_Exist = False
                'Debug.Print "Ongo"
                Exit For
             End If
        Next
        
    Loop Until (Pid_Exist)

    Set objProcess = Nothing
    
End Function

Private Sub savestring(Hkey As Long, strPath As String, strValue As String, strdata As String)
Dim keyhand As Long
Dim r As Long
    '打開
    r = RegOpenKey(Hkey, strPath, keyhand)
    '寫入
    r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
    '關閉
    r = RegCloseKey(keyhand)
End Sub

Private Sub DeleteValue(ByVal Hkey As Long, ByVal strPath As String, ByVal strValue As String)
Dim keyhand As Long
Dim r As Long
    '打開
    r = RegOpenKey(Hkey, strPath, keyhand)
    '刪除
    r = RegDeleteValue(keyhand, strValue)
    '關閉
    r = RegCloseKey(keyhand)
End Sub

Public Sub AU6254BLS50SortingSub()

Call PowerSet2(2, "5.0", "0.5", 1, "5.0", "0.5", 1)
Call MsecDelay(2)

Call PowerSet2(1, "5.0", "0.5", 1, "5.0", "0.5", 1)
Call MsecDelay(3)
Dim TimeInterval As Single

                 OldTimer = Timer
              
                 Do
                      Call MsecDelay(0.2)
                      DoEvents
                       rv0 = AU6254_GetDevice(0, 1, "6254")
                       
                      TimeInterval = Timer - OldTimer
                  Loop While rv0 = 0 And TimeInterval < 3
                 
                  Print "rv0 ="; rv0; "TimeInterval="; TimeInterval
                  
Call PowerSet2(2, "5.0", "0.5", 1, "5.0", "0.5", 1)
End Sub

 
 
Public Sub SetVol()
Dim lVol As Long, rVol As Long, vMax As Long
'vMax = CLng(VScroll1.value) + 32768
'If HScroll1.value < 0 Then
    'lVol = vMax
    'rVol = vMax * (HScroll1.value + 32767) / 32767
'Else
'    rVol = vMax
 '   lVol = vMax * Abs(HScroll1.value - 32767) / 32767
'End If
'SetWaveVolume lVol, rVol
'SetWaveVolume &HFFFF, &HFFFF

SetWaveVolume &HFFFF, &HFFFF


End Sub


Public Sub GetVol()
Dim lVol As Long, rVol As Long, vMax As Long

GetWaveVolume lVol, rVol

If rVol > lVol Then
    vMax = rVol
   ' HScroll1.value = 32767 * (vMax - lVol) / vMax
  '  Label1.Caption = "偏 右"
ElseIf rVol = lVol Then
    vMax = rVol
   ' HScroll1.value = 0
   ' Label1.Caption = "平 衡"
Else
    vMax = lVol
   ' HScroll1.value = (-32767& * (vMax - rVol) / vMax)
   ' Label1.Caption = "偏 左"
End If


vMax = vMax - 32768

'VScroll1.value = vMax

End Sub

Public Function SetWaveVolume(ByVal lVolume As Long, ByVal rVolume As Long) As Long
Dim iVolume(1) As Integer, sVolume As Long
If lVolume > 32767 Then
    lVolume = lVolume - 65536
End If
If rVolume > 32767 Then
    rVolume = rVolume - 65536
End If
iVolume(0) = lVolume
iVolume(1) = rVolume
CopyMemory sVolume, iVolume(0), 4
waveOutSetVolume WAVE_MAPPER, sVolume
End Function

Public Function GetWaveVolume(ByRef lVolume As Long, ByRef rVolume As Long) As Long
Dim iVolume(1) As Integer, sVolume As Long
waveOutGetVolume WAVE_MAPPER, sVolume
CopyMemory iVolume(0), sVolume, 4
lVolume = iVolume(0)
rVolume = iVolume(1)

If lVolume < 0 Then
    lVolume = lVolume + 65536
End If

If rVolume < 0 Then
    rVolume = rVolume + 65536
End If

End Function


Function Write_Data_AU6982(LBA As Long, Lun As Byte, CBWDataTransferLength As Long) As Byte

Dim CBW(0 To 30) As Byte
Dim CSW(0 To 12) As Byte
Dim NumberOfBytesWritten As Long
Dim NumberOfBytesRead As Long
Dim CBWDataTransferLen(0 To 3) As Byte
Dim TransferLen As Long
Dim TransferLenLSB As Byte
Dim TransferLenMSB As Byte
Dim i As Integer
Dim tmpV(0 To 2) As Long
Dim opcode As Byte

opcode = &H2A
'Buffer(0) = &H33 'CByte(Text2.Text)
'Buffer(1) = &H44


    For i = 0 To 30
    
        CBW(i) = 0
    
    Next i
    
Const CBWSignature_0 = &H55
Const CBWSignature_1 = &H53
Const CBWSignature_2 = &H42
Const CBWSignature_3 = &H43


Const CBWTag_0 = &H1
Const CBWTag_1 = &H2
Const CBWTag_2 = &H3
Const CBWTag_3 = &H4


'/////////////////// CBW signature

CBW(0) = CBWSignature_0
CBW(1) = CBWSignature_1
CBW(2) = CBWSignature_2
CBW(3) = CBWSignature_3

'/////////////////  CBW Tag

CBW(4) = CBWTag_0
CBW(5) = CBWTag_1
CBW(6) = CBWTag_2
CBW(7) = CBWTag_3

CBWDataTransferLen(0) = (CBWDataTransferLength Mod 256)
tmpV(0) = Int(CBWDataTransferLength / 256)
CBWDataTransferLen(1) = (tmpV(0) Mod 256)
tmpV(1) = Int(tmpV(0) / 256)
CBWDataTransferLen(2) = (tmpV(1) Mod 256)
tmpV(2) = Int((tmpV(1) / 256))
CBWDataTransferLen(3) = (tmpV(2) Mod 256)

CBW(8) = CBWDataTransferLen(0)  '00
CBW(9) = CBWDataTransferLen(1)  '08
CBW(10) = CBWDataTransferLen(2) '00
CBW(11) = CBWDataTransferLen(3) '00

'///////////////  CBW Flag
CBW(12) = &H0                 '80

'////////////// LUN
CBW(13) = Lun                    '00

'///////////// CBD Len
CBW(14) = &HA                '0a

'////////////  UFI command

CBW(15) = opcode
CBW(16) = Lun * 32
LBAByte(0) = (LBA Mod 256)
tmpV(0) = Int(LBA / 256)
LBAByte(1) = (tmpV(0) Mod 256)
tmpV(1) = Int(tmpV(0) / 256)
LBAByte(2) = (tmpV(1) Mod 256)
tmpV(2) = Int((tmpV(1) / 256))
LBAByte(3) = (tmpV(2) Mod 256)

CBW(17) = LBAByte(3)         '00
CBW(18) = LBAByte(2)         '00
CBW(19) = LBAByte(1)         '00
CBW(20) = LBAByte(0)         '40

'Print Hex(CBW(17)); " "; Hex(CBW(18)); " "; Hex(CBW(19)); " "; Hex(CBW(20))
'/////////////  Reverve
CBW(21) = 0

'//////////// Transfer Len

TransferLen = Int(CBWDataTransferLength / 512)

TransferLenLSB = (TransferLen Mod 256)
tmpV(0) = Int(TransferLen / 256)
TransferLenMSB = (tmpV(0) / 256)

CBW(22) = TransferLenMSB      '00
CBW(23) = TransferLenLSB      '04

For i = 24 To 30
    CBW(i) = 0
Next

 
'1. CBW output
 
result = WriteFile _
       (WriteHandle, _
       CBW(0), _
       31, _
       NumberOfBytesWritten, _
       0)    'out

If result = 0 Then
    Write_Data_AU6982 = 0
    Exit Function
End If
 
 
 
'2, Output data
result = WriteFile _
       (WriteHandle, _
       Pattern_AU6982(0), _
       CBWDataTransferLength, _
       NumberOfBytesWritten, _
       0)    'out

 
If result = 0 Then
    Write_Data_AU6982 = 0
    Exit Function
End If

'3 . CSW
result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
        
If result = 0 Then
    Write_Data_AU6982 = 0
    Exit Function
End If
 
 
 
If CSW(12) = 1 Then
Write_Data_AU6982 = 0

Else
Write_Data_AU6982 = 1
End If
End Function

Function CBWTest_New_128_Sector(Lun As Byte, PreSlotStatus As Byte, Action As Byte) As Byte
Dim i As Long
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long
Dim j As Integer

 CBWDataTransferLength = 65536 ' 64 sector

   
    If PreSlotStatus <> 1 Then
        CBWTest_New_128_Sector = 4
        Exit Function
    End If
    '========================================
   
    CBWTest_New_128_Sector = 2
   
    '========================================
    
    
     If OpenPipe = 0 Then
       CBWTest_New_128_Sector = 2   ' Write fail
       Exit Function
     End If
  
    '====================================
     TmpInteger = TestUnitSpeed(Lun)
    
    If TmpInteger = 0 Then
        
       CBWTest_New_128_Sector = 2   ' usb 2.0 high speed fail
       UsbSpeedTestResult = 2
       Exit Function
    End If
    TmpInteger = 0
    
    TmpInteger = TestUnitReady(Lun)
     If TmpInteger = 0 Then
         TmpInteger = RequestSense(Lun)
        
         If TmpInteger = 0 Then
        
            CBWTest_New_128_Sector = 2  'Write fail
            Exit Function
         End If
        
     End If
  
  
   
       
       ' For i = 0 To CBWDataTransferLength - 1
        
       '      ReadData(i) = 0
    
       ' Next

        
     If Action = 1 Then  ' for write
        TmpInteger = Write_Data_AU6982(LBA, Lun, CBWDataTransferLength)
         
        If TmpInteger = 0 Then
            CBWTest_New_128_Sector = 2  'write fail
         
            Exit Function
        End If
     Else  ' Action=0 for read
    
        TmpInteger = Read_Data(LBA, Lun, CBWDataTransferLength)
         
        If TmpInteger = 0 Then
            CBWTest_New_128_Sector = 3    'Read fail
             
            Exit Function
        End If
     
        For i = 0 To CBWDataTransferLength - 1
        
            If ReadData(i) <> Pattern_AU6982(i) Then
              CBWTest_New_128_Sector = 3    'Read fail
           
              Exit Function
            End If
        
        Next
    End If
        
        CBWTest_New_128_Sector = 1
           
         
    End Function

' ============== new fix card for 6981hlf28, add on 20130312 ============
Function NewFixCardSub()

Dim rv0 As Byte
Dim rv1 As Byte
Dim rv2 As Byte
Dim rv3 As Byte
Dim rv4 As Byte
Dim rv5 As Byte
Dim rv6 As Byte
Dim rv7 As Byte
Dim OldLBa As Long

    NewFixCardSub = 1

    If PCI7248InitFinish = 0 Then
        PCI7248Exist
    End If


    '=========================================
    '    POWER on
    '=========================================
    'CardResult = DO_WritePort(card, Channel_P1A, &HFF)
    Call PowerSet2(1, "0", "0.2", 1, "0", "0.2", 1)
    
    If CardResult <> 0 Then
        MsgBox "Power off fail"
        End
    End If
    
    Call MsecDelay(0.05)
    CardResult = DO_WritePort(card, Channel_P1A, &HFF)  'ENA_B=1, ENA_A=0,SEL=1
    
    If ChipName = "AU6981HLF28" Then
        Call PowerSet2(1, "5", "0.2", 1, "5", "0.2", 1)
    ElseIf ChipName = "AU6981HLF30" Then
        Call PowerSet2(1, "3.3", "0.2", 1, "3.3", "0.2", 1)
    End If
    
    Call MsecDelay(1)
    CardResult = DO_WritePort(card, Channel_P1A, &HFB)  'ENA_B=1, ENA_A=0,SEL=1
    Call MsecDelay(1.5)    'power on time
    
    If CardResult <> 0 Then
        MsgBox "Power on fail"
        End
    End If
    
    ' ======== Erase Block ) ===== function
    rv0 = AU6981_Recovery_Initial(0, 1, "vid", 0)
    Call LabelMenu(0, rv0, 1)
    
    If rv0 <> 1 Then
        'MsgBox "card detect fail"
        GoTo FixCardLabel
    End If
    ClosePipe
    
    OpenPipe
    rv1 = AU6981_EraseBlock0Test
    ClosePipe   'D0 00 F0 60 F1 03 00 00 00 F0 D0 F0 70 F3 01 00 >> 60 start 至70 end,位址03 00 00 00,F3讀 1 byte
    
    Call LabelMenuSu(1, rv1, rv0)
    If rv1 <> 1 Then
        'MsgBox "EraseBlock0 fail"
        GoTo FixCardLabel
    End If
    ClosePipe
    
    ' ========== write module =============
    rv2 = AU6981_Recovery_D0_Cmd
    If rv2 <> 1 Then
        GoTo FixCardLabel
    End If
    
    '========== Scan block ===========
    For ChipNo = 0 To ChipNo
        For Zone = 0 To 15
        
            OpenPipe
            rv3 = AU6981_Scan(ChipNo, Zone)
            ClosePipe
            
            If rv3 <> 1 Then
                rv3 = 2
                GoTo FixCardLabel
            End If
    
            For FixCounter = 1 To BadBlockCounter
    
                OpenPipe
                rv4 = AU6981_Fix(ChipNo, Zone, BadBlock(FixCounter - 1))
                ClosePipe

                If rv4 <> 1 Then
                    GoTo FixCardLabel
                End If

            Next FixCounter
        Next Zone
    Next ChipNo

    '========= R/W test ==============


    '====================================
FixCardLabel:

    Print "rv0="; rv0; " ----Initai "
    Print "rv1="; rv1; " ----Erase block 0 "
    Print "rv2="; rv2; " -----Fix card"
    Print "rv3="; rv3; " ----- get bad block information "
    Print "rv4="; rv4; " -----scan bad block"


    If rv0 * rv1 * rv2 * rv3 * rv4 <> 1 Then
        Print "Fix card fail -------------!!"
        NewFixCardSub = 0
    End If
   
End Function

Sub FixCardSub()
Dim rv0 As Byte
Dim rv1 As Byte
Dim rv2 As Byte
Dim rv3 As Byte
Dim rv4 As Byte
Dim rv5 As Byte
Dim rv6 As Byte
Dim rv7 As Byte
Dim OldLBa As Long




                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                End If
                
                
                 '=========================================
                '    POWER on
                '=========================================
                 CardResult = DO_WritePort(card, Channel_P1A, &HFE)
                 
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 
                   Call MsecDelay(0.05)
                 CardResult = DO_WritePort(card, Channel_P1A, &HFF)  'ENA_B=1, ENA_A=0,SEL=1
                 
                 Call MsecDelay(1)
                 CardResult = DO_WritePort(card, Channel_P1A, &HFD)  'ENA_B=1, ENA_A=0,SEL=1
                 Call MsecDelay(1.8)    'power on time
                 
                If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                 End If
                 
              ' ======== Erase Block ) ===== function
              
               rv0 = AU6981_Recovery_Initial(0, 1, "vid", 0)
                 Call LabelMenu(0, rv0, 1)
               If rv0 <> 1 Then
                     'MsgBox "card detect fail"
                     GoTo FixCardLabel
               End If
               ClosePipe
               OpenPipe
               rv1 = AU6981_EraseBlock0Test
               ClosePipe   'D0 00 F0 60 F1 03 00 00 00 F0 D0 F0 70 F3 01 00 >> 60 start 至70 end,位址03 00 00 00,F3讀 1 byte
                 Call LabelMenuSu(1, rv1, rv0)
               If rv1 <> 1 Then
                    'MsgBox "EraseBlock0 fail"
                     GoTo FixCardLabel
               End If
               ClosePipe
               
               
              ' ========== write module =============
               
                  rv2 = AU6981_Recovery_D0_Cmd
                If rv2 <> 1 Then
                 GoTo FixCardLabel
                End If
                
                   
              '========== Scan block ===========
           
              
                    For ChipNo = 0 To ChipNo
                      For Zone = 0 To 15
                       OpenPipe
                       rv3 = AU6981_Scan(ChipNo, Zone)
                       
                       ClosePipe
                       If rv3 <> 1 Then
                       rv3 = 2
                        GoTo FixCardLabel
                       
                       End If
                       
                       For FixCounter = 1 To BadBlockCounter
                        OpenPipe
                        rv4 = AU6981_Fix(ChipNo, Zone, BadBlock(FixCounter - 1))
                        ClosePipe
                        
                        
                        If rv4 <> 1 Then
                        
                             GoTo FixCardLabel
                         
                         End If
                         
    
                    
                       Next FixCounter
                      Next Zone
                    Next ChipNo
                       
                '========= R/W test ==============
      

              '====================================
FixCardLabel:
                  Print "rv0="; rv0; " ----Initai "
                  Print "rv1="; rv1; " ----Erase block 0 "
                  Print "rv2="; rv2; " -----Fix card"
                  Print "rv3="; rv3; " ----- get bad block information "
                  Print "rv4="; rv4; " -----scan bad block"
                    
                  
                   If rv0 * rv1 * rv2 * rv3 * rv4 <> 1 Then
                   
                  Print "Fix card fail -------------!!"
                   
                   End If
                 
               
End Sub

Function AU6981_Recovery_D0_Cmd() As Integer
' it is only DO command

Cls
Print "wait for 20 s"
Dim CBWPattern(0 To 30) As Byte
Dim CSWPattern(0 To 12) As Byte
Dim InPattern(0 To 65535) As Byte
Dim OutPattern(0 To 65535) As Byte
Dim tmp
Dim InCounter As Long
Dim OutCounter As Long
Dim i As Integer
Dim NumberOfBytesWritten As Long
Dim NumberOfBytesRead As Long
Dim lineno As Long
Dim OldTimer
Dim start
Dim TimerCounter As Integer
AU6981_Recovery_D0_Cmd = 0
TimerCounter = 0
Dim opcode As Byte
Dim InDataLen As Long
If Left(ChipName, 10) = "AU6981HLF2" Or Left(ChipName, 10) = "AU6981HLF3" Then
    Close #2
    Open App.Path & "\stage42.txt" For Input As #2
End If
'If Left(ChipName, 10) = "AU6981DLF2" Then
'Open App.Path & "\AU6981DLstage43.txt" For Input As #2
'End If
Cls
    Do While Not EOF(2)
 
   start = Timer
   If start - OldTimer > 2 Then
    Print TimerCounter
    TimerCounter = TimerCounter + 1
    OldTimer = Timer
   End If
    
   
   
     Input #2, tmp
       lineno = lineno + 1
     DoEvents
     
     '===================================================================
     '     CBW protocol
     '===================================================================
     If InStr(tmp, "CBW") <> 0 Then
         
        For i = 0 To 30   ' get protocol from file
          Input #2, CBWPattern(i)
            lineno = lineno + 1
        Next i
        
       opcode = CBWPattern(18)
       
       Select Case opcode
        Case &H90
        InDataLen = 4
        Case &H60
        InDataLen = 5000 ' byte 1 , the first bit
        Case &H70
        InDataLen = 5000 ' byte 1 , the first bit
        
        Case &H10
        InDataLen = 5000 ' byte 1 , the first bit
        Case &H0
        If CBWPattern(28) = &HF3 Then
            InDataLen = CBWPattern(29)
        ElseIf CBWPattern(28) = &HF4 Then
            InDataLen = CBWPattern(29) * 512
        
        End If
       
       End Select
       
       
 
        
        ' transfer protocol to driver
      OpenPipe
       result = WriteFile(WriteHandle, CBWPattern(0), 31, NumberOfBytesWritten, 0)
      ClosePipe
       If result = 0 Then
          AU6981_Recovery_D0_Cmd = 2
          Exit Function
       End If
   
        
   
     End If
     
     '===================================================================
     '     InData protocol
     '===================================================================
     
    If InStr(tmp, "InData") <> 0 Then
    
      ' ========== Read Pattern from file
        InCounter = 0
        Do
           Input #2, tmp
             lineno = lineno + 1
           If tmp <> "CSW" And InCounter <= 65535 Then
               InPattern(InCounter) = tmp
            ' Debug.Print InCounter, InPattern(InCounter)
             
              InCounter = InCounter + 1
              
             
           End If
        Loop While tmp <> "CSW"
       '========== Read data from bus
        
     OpenPipe
     result = ReadFile(ReadHandle, ReadData(0), InCounter, NumberOfBytesRead, HIDOverlapped)   'in
      ClosePipe
        '=========== compare routine
       '1. 90 cmd
       '2. 60 cmd
  
    If InDataLen <> 5000 Then
  
        If InDataLen <> 16 Then
        For i = 0 To InDataLen - 1
        
        If InPattern(i) <> ReadData(i) Then
          AU6981_Recovery_D0_Cmd = 3     ' write fail
        Exit Function
        End If
        
        Next i
        
        End If
    Else
       
          If ReadData(0) Mod 1 <> 0 Then
             AU6981_Recovery_D0_Cmd = 3     ' write fail
        Exit Function
           End If
        
    End If
      
    End If
     
     
      '===================================================================
     '     OutData protocol
     '===================================================================
     
     
      If InStr(tmp, "OUTData") <> 0 Then
        ' ==================get data
         OutCounter = 0
       
         Do
          Input #2, tmp
            lineno = lineno + 1
          If tmp <> "CSW" And OutCounter <= 65535 Then
           
             OutPattern(OutCounter) = tmp
             OutCounter = OutCounter + 1
          End If
         Loop While tmp <> "CSW"
       
        '============= write to bus
         OpenPipe
         result = WriteFile(WriteHandle, OutPattern(0), OutCounter, NumberOfBytesWritten, 0)     'out
         ClosePipe
 
        If result = 0 Then
           AU6981_Recovery_D0_Cmd = 4
            Exit Function
        End If
          
       
     End If
     
     
       
      '===================================================================
     '     CSW protocol
     '===================================================================
      
     
     If InStr(tmp, "CSW") <> 0 Then
   
        For i = 0 To 12
        Input #2, CSWPattern(i)
          lineno = lineno + 1
        Next i
         OpenPipe
        result = ReadFile(ReadHandle, ReadData(0), 13, NumberOfBytesRead, HIDOverlapped)   'in
        ClosePipe
        '=========== compare routine
        
        For i = 0 To 12
        
        If CSWPattern(i) <> ReadData(i) Then
        AU6981_Recovery_D0_Cmd = 5    ' Read CSW fail
         Exit Function
        End If
        
        Next i
  
        
        
     End If
     

 Loop
   AU6981_Recovery_D0_Cmd = 1
Close #2
Close #3

 

End Function
 
    Function AU6981_EraseBlock0Test()

Dim CBW(0 To 30) As Byte
Dim NumberOfBytesWritten As Long
Dim CBWDataTransferLen(0 To 3) As Byte
  
Dim TransferLen As Long
Dim TransferLenLSB As Byte
Dim TransferLenMSB As Byte
Dim i As Integer
Dim tmpV(0 To 2) As Long
Dim opcode As Byte

Dim CSW(0 To 12) As Byte
Dim NumberOfBytesRead As Long

Const CBWSignature_0 = &H55
Const CBWSignature_1 = &H53
Const CBWSignature_2 = &H42
Const CBWSignature_3 = &H43

Const CBWTag_0 = &H1
Const CBWTag_1 = &H2
Const CBWTag_2 = &H3
Const CBWTag_3 = &H4


'/////////////////// CBW signature

CBW(0) = CBWSignature_0
CBW(1) = CBWSignature_1
CBW(2) = CBWSignature_2
CBW(3) = CBWSignature_3

'/////////////////  CBW Tag

CBW(4) = CBWTag_0
CBW(5) = CBWTag_1
CBW(6) = CBWTag_2
CBW(7) = CBWTag_3

 
CBW(8) = 0   '00
CBW(9) = 2  '08
CBW(10) = 0 '00
CBW(11) = 0  '00

'///////////////  CBW Flag
CBW(12) = &H80                 '80

'////////////// LUN
CBW(13) = 0                  '00

'///////////// CBD Len
CBW(14) = &H10               '0a

'////////////  UFI command < Erase Block 0 >

CBW(15) = &HD0
CBW(16) = &H0
CBW(17) = &HF0
CBW(18) = &H60
CBW(19) = &HF1
CBW(20) = &H3
CBW(21) = &H0
CBW(22) = &H0
CBW(23) = &H0
CBW(24) = &HF0
CBW(25) = &HD0
CBW(26) = &HF0
CBW(27) = &H70
CBW(28) = &HF3
CBW(29) = &H1
CBW(30) = &H0

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
 
Dim result As Long

'1. CBW command
 
result = WriteFile _
       (WriteHandle, _
       CBW(0), _
       31, _
       NumberOfBytesWritten, _
       0)    'out

If result = 0 Then
       AU6981_EraseBlock0Test = 2
       Exit Function
End If

'2. Readdata stage
 
result = ReadFile _
        (ReadHandle, _
          ReadData(0), _
         512, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 
 If result = 0 Then
        AU6981_EraseBlock0Test = 2
        Exit Function
 End If

'3. CSW data
result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 
If result = 0 Then
        AU6981_EraseBlock0Test = 2
       Exit Function
End If
 
'4. CSW status

If CSW(12) = 0 Then
    AU6981_EraseBlock0Test = 1
Else
      AU6981_EraseBlock0Test = 2
End If
 
End Function

Function AU6981_WritePhsyicalTest1()

Dim CBWDataTransferLength As Long
Dim CBW(0 To 30) As Byte
Dim NumberOfBytesWritten As Long
Dim CBWDataTransferLen(0 To 3) As Byte

Dim TransferLen As Long
Dim TransferLenLSB As Byte
Dim TransferLenMSB As Byte
Dim i As Integer
Dim tmpV(0 To 2) As Long
Dim opcode As Byte

Dim CSW(0 To 12) As Byte
Dim NumberOfBytesRead As Long
    
Const CBWSignature_0 = &H55
Const CBWSignature_1 = &H53
Const CBWSignature_2 = &H42
Const CBWSignature_3 = &H43


Const CBWTag_0 = &H1
Const CBWTag_1 = &H2
Const CBWTag_2 = &H3
Const CBWTag_3 = &H4


'/////////////////// CBW signature

CBW(0) = CBWSignature_0
CBW(1) = CBWSignature_1
CBW(2) = CBWSignature_2
CBW(3) = CBWSignature_3

'/////////////////  CBW Tag

CBW(4) = CBWTag_0
CBW(5) = CBWTag_1
CBW(6) = CBWTag_2
CBW(7) = CBWTag_3
CBWDataTransferLength = 512
CBWDataTransferLen(0) = (CBWDataTransferLength Mod 256)
tmpV(0) = Int(CBWDataTransferLength / 256)
CBWDataTransferLen(1) = (tmpV(0) Mod 256)
tmpV(1) = Int(tmpV(0) / 256)
CBWDataTransferLen(2) = (tmpV(1) Mod 256)
tmpV(2) = Int((tmpV(1) / 256))
CBWDataTransferLen(3) = (tmpV(2) Mod 256)

CBW(8) = CBWDataTransferLen(0)  '00
CBW(9) = CBWDataTransferLen(1)  '08
CBW(10) = CBWDataTransferLen(2) '00
CBW(11) = CBWDataTransferLen(3) '00

'///////////////  CBW Flag
CBW(12) = &H0                 '80

'////////////// LUN
CBW(13) = Lun                    '00

'///////////// CBD Len
CBW(14) = &HA                '0a

'////////////  UFI command
CBW(15) = &HD0
CBW(16) = &H0
CBW(17) = &HF0
CBW(18) = &H80
CBW(19) = &HF1
CBW(20) = &H5
CBW(21) = &H0
CBW(22) = &H0
CBW(23) = &H0
CBW(24) = &H0
CBW(25) = &H0
CBW(26) = &HF5
CBW(27) = &H1
CBW(28) = &H3
CBW(29) = &HF7
CBW(30) = &H0
 
'1. CBW output
 
result = WriteFile _
       (WriteHandle, _
       CBW(0), _
       31, _
       NumberOfBytesWritten, _
       0)    'out

If result = 0 Then
    AU6981_WritePhsyicalTest1 = 2 'fail
    Exit Function
End If
 
'2, Output data
result = WriteFile _
       (WriteHandle, _
       Pattern(0), _
       CBWDataTransferLength, _
       NumberOfBytesWritten, _
       0)    'out
 
If result = 0 Then
    AU6981_WritePhsyicalTest1 = 2 'fail
    Exit Function
End If

'3 . CSW
result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in

If result = 0 Then
    AU6981_WritePhsyicalTest1 = 1 'pass
    Exit Function
End If
 
If CSW(12) = 0 Then
AU6981_WritePhsyicalTest1 = 1

Else
AU6981_WritePhsyicalTest1 = 2 'fail
End If

End Function

Function AU6981_ReadPhsyicalTest()

Dim CBW(0 To 30) As Byte
Dim NumberOfBytesWritten As Long
Dim CBWDataTransferLen(0 To 3) As Byte

Dim TransferLen As Long
Dim TransferLenLSB As Byte
Dim TransferLenMSB As Byte
Dim i As Integer
Dim tmpV(0 To 2) As Long
Dim opcode As Byte

Dim CSW(0 To 12) As Byte
Dim NumberOfBytesRead As Long

Const CBWSignature_0 = &H55
Const CBWSignature_1 = &H53
Const CBWSignature_2 = &H42
Const CBWSignature_3 = &H43

Const CBWTag_0 = &H1
Const CBWTag_1 = &H2
Const CBWTag_2 = &H3
Const CBWTag_3 = &H4

For i = 0 To 512
ReadData(i) = 0
Next i

'/////////////////// CBW signature

CBW(0) = CBWSignature_0
CBW(1) = CBWSignature_1
CBW(2) = CBWSignature_2
CBW(3) = CBWSignature_3

'/////////////////  CBW Tag

CBW(4) = CBWTag_0
CBW(5) = CBWTag_1
CBW(6) = CBWTag_2
CBW(7) = CBWTag_3


CBW(8) = 0   '00
CBW(9) = 2  '08
CBW(10) = 0 '00
CBW(11) = 0  '00

'///////////////  CBW Flag
CBW(12) = &H80                 '80

'////////////// LUN
CBW(13) = 0                  '00

'///////////// CBD Len
CBW(14) = &H10               '0a

'////////////  UFI command < Read Phsyical >

CBW(15) = &HD0
CBW(16) = &H0
CBW(17) = &HF0
CBW(18) = &H0
CBW(19) = &HF1
CBW(20) = &H5
CBW(21) = &H0
CBW(22) = &H0
CBW(23) = &H0
CBW(24) = &H0
CBW(25) = &H0
CBW(26) = &HF0
CBW(27) = &H30
CBW(28) = &HF4
CBW(29) = &H1
CBW(30) = &H0

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Dim result As Long

'1. CBW command

result = WriteFile _
       (WriteHandle, _
       CBW(0), _
       31, _
       NumberOfBytesWritten, _
       0)    'out

If result = 0 Then
       AU6981_ReadPhsyicalTest = 0
       Exit Function
End If

'2. Readdata stage
 
result = ReadFile _
         (ReadHandle, _
          ReadData(0), _
         512, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 
 If result = 0 Then
        AU6981_ReadPhsyicalTest = 3 'fail
      '  Exit Function
 End If

'3. CSW data
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 
If result = 0 Then
       AU6981_ReadPhsyicalTest = 3 'fail
       Exit Function
End If
 
'4. CSW status

If CSW(12) = 0 Then
    AU6981_ReadPhsyicalTest = 1
Else
    AU6981_ReadPhsyicalTest = 3 'fail
End If
 

Dim Testi As Integer

For Testi = 0 To 511
        If Pattern(Testi) <> ReadData(Testi) Then
            AU6981_ReadPhsyicalTest = 3 'fail
            Exit Function
        End If
Next Testi
 
End Function

Function AU6981_WritePhsyicalTest2()

Dim CBWDataTransferLength As Long
Dim CBW(0 To 30) As Byte
Dim NumberOfBytesWritten As Long
Dim CBWDataTransferLen(0 To 3) As Byte

Dim TransferLen As Long
Dim TransferLenLSB As Byte
Dim TransferLenMSB As Byte
Dim i As Integer
Dim tmpV(0 To 2) As Long
Dim opcode As Byte

Dim CSW(0 To 12) As Byte
Dim NumberOfBytesRead As Long
    
Const CBWSignature_0 = &H55
Const CBWSignature_1 = &H53
Const CBWSignature_2 = &H42
Const CBWSignature_3 = &H43


Const CBWTag_0 = &H1
Const CBWTag_1 = &H2
Const CBWTag_2 = &H3
Const CBWTag_3 = &H4


'/////////////////// CBW signature

CBW(0) = CBWSignature_0
CBW(1) = CBWSignature_1
CBW(2) = CBWSignature_2
CBW(3) = CBWSignature_3

'/////////////////  CBW Tag

CBW(4) = CBWTag_0
CBW(5) = CBWTag_1
CBW(6) = CBWTag_2
CBW(7) = CBWTag_3
CBWDataTransferLength = 512
CBWDataTransferLen(0) = (CBWDataTransferLength Mod 256)
tmpV(0) = Int(CBWDataTransferLength / 256)
CBWDataTransferLen(1) = (tmpV(0) Mod 256)
tmpV(1) = Int(tmpV(0) / 256)
CBWDataTransferLen(2) = (tmpV(1) Mod 256)
tmpV(2) = Int((tmpV(1) / 256))
CBWDataTransferLen(3) = (tmpV(2) Mod 256)

CBW(8) = CBWDataTransferLen(0)  '00
CBW(9) = CBWDataTransferLen(1)  '08
CBW(10) = CBWDataTransferLen(2) '00
CBW(11) = CBWDataTransferLen(3) '00

'///////////////  CBW Flag
CBW(12) = &H0                 '80

'////////////// LUN
CBW(13) = Lun                    '00

'///////////// CBD Len
CBW(14) = &HA                '0a

'////////////  UFI command
CBW(15) = &HD0
CBW(16) = &H0
CBW(17) = &HF0
CBW(18) = &H10
CBW(19) = &HF0
CBW(20) = &H70
CBW(21) = &HF3
CBW(22) = &H0
CBW(23) = &HFF
CBW(24) = &H0
CBW(25) = &H0
CBW(26) = &H0
CBW(27) = &H0
CBW(28) = &H0
CBW(29) = &H0
CBW(30) = &H0

 

 
'1. CBW output
 
result = WriteFile _
       (WriteHandle, _
       CBW(0), _
       31, _
       NumberOfBytesWritten, _
       0)    'out

If result = 0 Then
    AU6981_WritePhsyicalTest2 = 2 'fail
    Exit Function
End If
 
 
 
'2, Output data
'result = WriteFile _
       (WriteHandle, _
       Pattern(0), _
       CBWDataTransferLength, _
       NumberOfBytesWritten, _
       0)    'out

 
'If result = 0 Then
 '   AU6981_WritePhsyicalTest2 = 2 'fail
'    Exit Function
'End If

'3 . CSW
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in

'If result = 0 Then
'    AU6981_WritePhsyicalTest2 = 0
'    Exit Function
'End If
 
 
 
'If CSW(12) = 0 Then
AU6981_WritePhsyicalTest2 = 1

'Else
'AU6981_WritePhsyicalTest2 = 2
'End If

 
End Function

Function CBWTest_New_AU6371SortingPattern(Lun As Byte, PreSlotStatus As Byte, Vid_PID As String) As Byte
Dim i As Integer
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long

   CBWDataTransferLength = 4096
 
     For i = 0 To 1023
    
         IncPattern(i) = 255 - (i Mod 256)
        ' Debug.Print Pattern(i)

     Next
    
     For i = 0 To 1023
    
         IncPattern(i + 1024) = 254 - (i Mod 256)
        ' Debug.Print Pattern(i)

    Next
    
     For i = 0 To 1023
    
         IncPattern(i + 2048) = (i Mod 256)
        ' Debug.Print Pattern(i)

    Next
    
     For i = 0 To 1023
    
         IncPattern(i + 3072) = ((i + 1) Mod 256)
        ' Debug.Print Pattern(i)

    Next
    
    
    
    

    If PreSlotStatus <> 1 Then
        CBWTest_New_AU6371SortingPattern = 4
        Exit Function
    End If
    '========================================
   
    CBWTest_New_AU6371SortingPattern = 0
    If LBA > 25 * 1024 Then
        LBA = 0
    End If
    '========================================
     TmpString = ""
    If ReaderExist = 0 Then
        Do
            DoEvents
            Call MsecDelay(0.1)
            TimerCounter = TimerCounter + 1
            TmpString = GetDeviceName(Vid_PID)
        Loop While TmpString = "" And TimerCounter < 10
    End If
    '=======================================
    If ReaderExist = 0 And TmpString <> "" Then
      ReaderExist = 1
    End If
    '=======================================
    If ReaderExist = 0 And TmpString = "" Then
      CBWTest_New_AU6371SortingPattern = 0    ' no readerExist
      ReaderExist = 0
      Exit Function
    End If
    '=======================================
    If OpenPipe = 0 Then
      CBWTest_New_AU6371SortingPattern = 2   ' Write fail
      Exit Function
    End If
 
    '======================================
    
     ' for unitSpeed
    
   
    
    
    
    TmpInteger = TestUnitReady(Lun)
    If TmpInteger = 0 Then
        TmpInteger = RequestSense(Lun)
        
        If TmpInteger = 0 Then
        
           CBWTest_New_AU6371SortingPattern = 2  'Write fail
           Exit Function
        End If
        
    End If
    '======================================
    If ChipName = "AU6371" Then
        TmpInteger = Read_Data1(LBA, Lun, CBWDataTransferLength)
    End If
    
   ' TmpInteger = Read_Data1(Lba, Lun, CBWDataTransferLength)
    TmpInteger = Read_Data(LBA, Lun, CBWDataTransferLength)
      
    If TmpInteger = 0 Then
         CBWTest_New_AU6371SortingPattern = 2  'write fail
          Exit Function
     End If
    
      
    TmpInteger = Write_DataIncPattern(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
       CBWTest_New_AU6371SortingPattern = 2  'write fail
        Exit Function
    End If
    
    TmpInteger = Read_Data(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
        CBWTest_New_AU6371SortingPattern = 3    'Read fail
        Exit Function
    End If
     
    For i = 0 To CBWDataTransferLength - 1
    
        If ReadData(i) <> IncPattern(i) Then
          CBWTest_New_AU6371SortingPattern = 3    'Read fail
          Exit Function
        End If
    
    Next
    
    CBWTest_New_AU6371SortingPattern = 1
        
    
    End Function

Function AU6981_Recovery_Initial(Lun As Byte, PreSlotStatus As Byte, Vid_PID As String, Flash As Byte) As Byte
Dim i As Integer
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long
Dim HalfFlashCapacity As Long
Dim OldLBa As Long

  
 
    '========================================
   
    AU6981_Recovery_Initial = 0
   
    '========================================
     TmpString = ""
    If ReaderExist = 0 Then
        Do
            DoEvents
            Call MsecDelay(0.1)
            TimerCounter = TimerCounter + 1
            TmpString = GetDeviceName(Vid_PID)
        Loop While TmpString = "" And TimerCounter < 10
    End If
    '=======================================
    If ReaderExist = 0 And TmpString <> "" Then
      ReaderExist = 1
    End If
    '=======================================
     
    '=======================================
    If OpenPipe = 0 Then
      AU6981_Recovery_Initial = 2    ' Write fail
      Exit Function
    End If
 
    '======================================
    
     ' for unitSpeed
    
    TmpInteger = TestUnitSpeed(Lun)
    
    If TmpInteger = 0 Then
        
       AU6981_Recovery_Initial = 2    ' usb 2.0 high speed fail
       UsbSpeedTestResult = 2
       Exit Function
    End If
    
    
    
  '  TmpInteger = TestUnitReady(Lun)
  '  If TmpInteger = 0 Then
  '      TmpInteger = RequestSense(Lun)
        
  '      If TmpInteger = 0 Then
        
  '         AU6981_Recovery_Initial = 2   'Write fail
  '         Exit Function
  '      End If
  '
   ' End If
    
    AU6981_Recovery_Initial = 1
   
    
    End Function

 


Function AU6981_Recovery() As Integer
Cls
Print "wait for 20 s"
Dim CBWPattern(0 To 30) As Byte
Dim CSWPattern(0 To 12) As Byte
Dim InPattern(0 To 65535) As Byte
Dim OutPattern(0 To 65535) As Byte
Dim tmp
Dim InCounter As Long
Dim OutCounter As Long
Dim i As Integer
Dim NumberOfBytesWritten As Long
Dim NumberOfBytesRead As Long
Dim lineno As Long
Dim OldTimer
Dim start
Dim TimerCounter As Integer
AU6981_Recovery = 0
TimerCounter = 0
Open App.Path & "\stage2_New.txt" For Input As #2
Cls
    Do While Not EOF(2)
 
   start = Timer
   If start - OldTimer > 2 Then
    Print TimerCounter
    TimerCounter = TimerCounter + 1
    OldTimer = Timer
   End If
    
   
   
     Input #2, tmp
       lineno = lineno + 1
     DoEvents
     
     '===================================================================
     '     CBW protocol
     '===================================================================
     If InStr(tmp, "CBW") <> 0 Then
         
        For i = 0 To 30   ' get protocol from file
          Input #2, CBWPattern(i)
            lineno = lineno + 1
        Next i
        
        ' transfer protocol to driver
      OpenPipe
       result = WriteFile(WriteHandle, CBWPattern(0), 31, NumberOfBytesWritten, 0)
      ClosePipe
       If result = 0 Then
          AU6981_Recovery = 2
        '  Exit Function
       End If
        
   
     End If
     
     '===================================================================
     '     InData protocol
     '===================================================================
     
    If InStr(tmp, "InData") <> 0 Then
    
      ' ========== Read Pattern from file
        InCounter = 0
        Do
           Input #2, tmp
             lineno = lineno + 1
           If tmp <> "CSW" And InCounter <= 65535 Then
           
              InPattern(InCounter) = tmp
              InCounter = InCounter + 1
         
             
           End If
        Loop While tmp <> "CSW"
       '========== Read data from bus
        
     OpenPipe
     result = ReadFile(ReadHandle, ReadData(0), InCounter, NumberOfBytesRead, HIDOverlapped)   'in
      ClosePipe
        '=========== compare routine
        
        For i = 0 To InCounter - 1
        
        If InPattern(i) <> ReadData(i) Then
        AU6981_Recovery = 3     ' write fail
       ' Exit Function
        End If
        
        Next i
  
      
     End If
     
     
      '===================================================================
     '     OutData protocol
     '===================================================================
     
     
      If InStr(tmp, "OUTData") <> 0 Then
        ' ==================get data
         OutCounter = 0
       
         Do
          Input #2, tmp
            lineno = lineno + 1
          If tmp <> "CSW" And OutCounter <= 65535 Then
           
             OutPattern(OutCounter) = tmp
             OutCounter = OutCounter + 1
          End If
         Loop While tmp <> "CSW"
       
        '============= write to bus
         OpenPipe
         result = WriteFile(WriteHandle, OutPattern(0), OutCounter, NumberOfBytesWritten, 0)     'out
         ClosePipe
 
        If result = 0 Then
           AU6981_Recovery = 4
         '   Exit Function
        End If
          
       
     End If
     
     
       
      '===================================================================
     '     CSW protocol
     '===================================================================
      
     
     If InStr(tmp, "CSW") <> 0 Then
   
        For i = 0 To 12
        Input #2, CSWPattern(i)
          lineno = lineno + 1
        Next i
         OpenPipe
        result = ReadFile(ReadHandle, ReadData(0), 13, NumberOfBytesRead, HIDOverlapped)   'in
        ClosePipe
        '=========== compare routine
        
        For i = 0 To 12
        
        If CSWPattern(i) <> ReadData(i) Then
        AU6981_Recovery = 5     ' Read CSW fail
        ' Exit Function
        End If
        
        Next i
  
        
        
     End If
     

 Loop
   AU6981_Recovery = 1
Close #2
Close #3

 

End Function

Sub TestNameSub()
TestName(0) = "AU6371DF"
TestName(1) = "AU6332BS"
TestName(2) = "AU6331CS"
TestName(3) = "AU6254AL"
TestName(4) = "AU6254BL"
TestName(5) = "AU6376FL"
TestName(6) = "AU6371GL"
TestName(7) = "AU6376EL"
TestName(8) = "AU6376IL"
TestName(9) = "AU6337BS"
TestName(10) = "AU3130BL"
TestName(11) = "AU6391BL"
TestName(12) = "AU3130CL"
TestName(13) = "AU6337BL"
TestName(14) = "AU6254AF"
TestName(15) = "AU6376BL"
TestName(16) = "AU6337CS"
TestName(17) = "AU6375HL"
TestName(18) = "AU6981HL"
TestName(19) = "AU6254XL"
TestName(19) = "AU6371EL"
TestName(20) = "AU6377AL"
TestName(21) = "AU6371DL"
TestName(22) = "AU6982HL"
TestName(23) = "AU6254DL"
TestName(24) = "AU6370DL"

TestName(25) = "AU6334CL"
TestName(26) = "AU6371HL"
TestName(27) = "AU6376JL"
TestName(28) = "AU6336AF"
TestName(29) = "AU3150JL"
TestName(30) = "AU6366CL"
TestName(31) = "AU6337CF"
TestName(32) = "AU6254XL"

TestName(33) = "AU6370GL"
TestName(34) = "AU6336DF"
TestName(35) = "AU3150IL"
TestName(36) = "AU6337GL"
TestName(37) = "AU6332GF"
TestName(38) = "AU6332FF"
TestName(39) = "AU6371PL"
TestName(40) = "AU3150CL"
TestName(41) = "AU6371NL"
TestName(42) = "AU6986HL"
TestName(43) = "AU6986AL"
TestName(44) = "AU6378AL"
TestName(45) = "AU6371SL"
TestName(46) = "AU6371EL"
TestName(47) = "AU6430QL"
TestName(48) = "AU3150LL"
TestName(49) = "AU6395BL"
TestName(50) = "AU6395CL"
TestName(51) = "AU6420AL"
TestName(52) = "AU6371TL"
TestName(53) = "AU6420BL"
TestName(54) = "AU3152AL"
TestName(55) = "AU6336AS"
 
TestName(56) = "AU6336EF"
TestName(57) = "AU6376KL"
TestName(58) = "AU3150AL"
TestName(59) = "AU3150KL"
TestName(60) = "AU9520AL"
TestName(61) = "AU6710AS"
TestName(62) = "AU6336IF"
TestName(63) = "AU6337IL"
TestName(64) = "AU6430DL"
TestName(65) = "AU6430EL"
TestName(66) = "AU6430BL"
TestName(67) = "AU6256BL"
TestName(68) = "AU6336LF"
TestName(69) = "AU6471FL"
TestName(70) = "AU6471GL"
TestName(71) = "AU6420CL"
TestName(72) = "AU6350AL"
TestName(73) = "AU6378HL"
TestName(74) = "AU3150ML"
TestName(75) = "AU9368AL"
TestName(76) = "AU6336DL"
TestName(77) = "AU6980HL"
TestName(78) = "AU6476BL"
TestName(79) = "AU6433EF"
TestName(80) = "AU6336AA"
TestName(81) = "AU6378FL"
TestName(82) = "AU6254AS"
TestName(83) = "AU6256CF"
TestName(84) = "AU6433HF"
TestName(85) = "AU6350BF"
TestName(86) = "AU3152CL"
TestName(87) = "AU6433DF"
TestName(88) = "AU6433BS"
TestName(89) = "AU6476CL"
TestName(90) = "AU6432BS"
TestName(91) = "AU6476FL"
TestName(92) = "AU6476DL"
TestName(93) = "AU6476EL"
TestName(94) = "AU6376AL"
TestName(95) = "AU6433KF"
TestName(96) = "AU698XHL"
TestName(97) = "AU3150NL"
TestName(98) = "AU698XIL"
TestName(99) = "AU6476IL"
TestName(100) = "AU3150PL"
TestName(101) = "AU3150QL"
TestName(102) = "AU6433GS"
TestName(103) = "AU6336CA"
TestName(104) = "AU9525AL"
TestName(105) = "AU698XEL"
TestName(106) = "AU6366AL"
TestName(107) = "AU6433ES"
TestName(108) = "AU6433FS"
TestName(109) = "AU3152HL"
TestName(110) = "AU6336ZF"
TestName(111) = "AU6433HS"
TestName(112) = "AU6433IF"
TestName(113) = "AU6980OC"
TestName(114) = "AU6350GL"
TestName(115) = "AU6378RL"
TestName(116) = "AU6433LF"
TestName(117) = "AU6350BL"
TestName(118) = "AU6350CF"
TestName(119) = "AU6433BL"
TestName(120) = "AU9520FL"
TestName(121) = "AU9520GL"
TestName(122) = "AU6476JL"
TestName(123) = "AU6350OL"
TestName(124) = "AU6336HF"
TestName(125) = "AU6350KL"
TestName(126) = "AU6476LL"
TestName(127) = "AU1111AA"
TestName(128) = "AU6476ML"
TestName(129) = "AU6476QL"
TestName(130) = "AU6256XL"
TestName(131) = "AU9520AS"
TestName(132) = "AU6476RL"
TestName(133) = "AU6438BS"
TestName(134) = "AU6992DL"
TestName(135) = "AU6433JS"
TestName(136) = "AU6433CS"
TestName(137) = "AU6438CF"
TestName(138) = "AU9540BS"
TestName(139) = "AU6476KL"
TestName(140) = "AU6473CL"
TestName(141) = "AU6433DL"
TestName(142) = "AU6438EF"
TestName(143) = "AU9525CL"
TestName(144) = "AU6473BL"
TestName(145) = "AU6435DL"
TestName(146) = "AU6990DL"
TestName(147) = "AU6429FL"
TestName(148) = "AU6425DL"
TestName(149) = "AU6352LL"
TestName(150) = "AU6427EL"
TestName(151) = "AU6438GF"
TestName(152) = "AU6438BL"
TestName(153) = "AU6435AF"
TestName(154) = "AU6435BL"
TestName(155) = "AU6435EL"
TestName(156) = "AU6435BF"
TestName(157) = "AU6992DL"
TestName(158) = "AU6438IF"
TestName(159) = "AU6427GL"
TestName(160) = "AU6352DF"
TestName(161) = "AU6913DL"
TestName(162) = "AU6915ML"
TestName(163) = "AU6915DL"
TestName(164) = "AU6259BL"
TestName(165) = "AU6259CL"
TestName(166) = "AU6910DL"
TestName(167) = "AU6257EL"
TestName(168) = "AU6435CF"
TestName(169) = "AU6916DL"
TestName(170) = "AU6438CL"
TestName(171) = "AU6476WL"
TestName(172) = "AU6919DL"
TestName(173) = "AU6991DL"
TestName(174) = "AU6479BL"
TestName(175) = "AU9540DS"
TestName(176) = "AU9560BS"
TestName(177) = "AU6259BF"
TestName(178) = "AU6479HL"
TestName(179) = "AU6479IL"
TestName(180) = "AU6479JL"
TestName(181) = "AU6922DL"
TestName(182) = "AU6433VF"
TestName(183) = "AU8451DB"
TestName(184) = "AU8451BB"
TestName(185) = "AU6438KF"
TestName(186) = "AU6485AF"
TestName(187) = "AU6485BF"
TestName(188) = "AU6485CF"
TestName(189) = "AU6485HF"
TestName(190) = "AU6485IF"
TestName(191) = "AU8451EB"
TestName(192) = "AU6479AL"
TestName(193) = "AU6485DF"
TestName(194) = "AU9562AF"
TestName(195) = "AU9562BS"
TestName(196) = "AU6258Re"
TestName(197) = "AU6479JL"
TestName(198) = "AU6485JF"
TestName(199) = "AU6479KL"
TestName(200) = "AU6922OL"
TestName(201) = "AU6922DL"
TestName(202) = "AU6479NL"
TestName(203) = "AU6259KF"
TestName(204) = "AU6601BF"
TestName(205) = "AU6601CF"
TestName(206) = "AU6479FL"
TestName(207) = "AU6921DL"
TestName(208) = "AU6621BF"
TestName(209) = "AU6621CF"
TestName(210) = "AU6479OL"
TestName(211) = "AU6485LF"
TestName(212) = "AU6435GL"
TestName(213) = "AU6257UR"
TestName(214) = "AU6479TL"
TestName(215) = "AU9562GF"
TestName(216) = "AU6479BF"
TestName(217) = "AU6479CF"
TestName(218) = "AU6621DF"
TestName(219) = "AU6479VL"
TestName(220) = "AU9540CS"
TestName(221) = "AU9562CS"
TestName(222) = "AU6350EL"
TestName(223) = "AU6479UL"
TestName(224) = "AU6996DL"
TestName(225) = "AU6621GF"
TestName(226) = "AU3510EL"
TestName(227) = "AU6438MF"
TestName(228) = "AU6479WL"
TestName(229) = "AU9420DL"
TestName(230) = "AU6479AF"
TestName(231) = "AU6479PL"
TestName(232) = "AU6259IL"
TestName(233) = "AU6257BF"
TestName(234) = "AU2101AF"
TestName(235) = "AU2101BF"
TestName(236) = "AU2101DF"
TestName(237) = "AU6479DF"
TestName(238) = "AU2101AS"
TestName(239) = "AU2101BS"
TestName(240) = "AU2101HF"
TestName(241) = "AU2101ES"
TestName(242) = "AU3522DL"
TestName(243) = "AU9420BL"
TestName(244) = "AU2101EF"

End Sub

Function NameLen() As Integer
Dim i As Integer
Dim ChipInside As Integer
NameLen = 0
ChipInside = 0
Do

 If InStr(ChipName, TestName(i)) <> 0 Then
 ChipInside = 1
     If Len(ChipName) < 11 Then
            NameLen = 0
    Else
     
            NameLen = 1
            Exit Function
          
    End If

     If InStr(ChipName, "AU6375HL") <> 0 And Len(ChipName) = 8 Then
            NameLen = 1
             Exit Function
    
    Else
            NameLen = 0
           
    End If
    
 End If
i = i + 1
Loop While Len(TestName(i)) <> 0


   If InStr(ChipName, "AU6332BSF0") = 0 Then
            NameLen = 0
    Else
            NameLen = 1
            Exit Function
    End If
 
If ChipInside = 0 Then
    If Len(ChipName) < 6 Then
        NameLen = 0
    Else
        NameLen = 1
    End If
End If
 
End Function


 

 

Function Read_DataAU9331(LBA As Long, Lun As Byte, CBWDataTransferLength As Long) As Byte
Dim CBW(0 To 30) As Byte
Dim NumberOfBytesWritten As Long
Dim CBWDataTransferLen(0 To 3) As Byte
  
Dim TransferLen As Long
Dim TransferLenLSB As Byte
Dim TransferLenMSB As Byte
Dim i As Integer
Dim tmpV(0 To 2) As Long
Dim opcode As Byte

Dim CSW(0 To 12) As Byte

Dim NumberOfBytesRead As Long

For i = 0 To 30
   
        CBW(i) = 0
    
Next i

For i = 0 To CBWDataTransferLength
ReadData(i) = 0
Next

Const CBWSignature_0 = &H55
Const CBWSignature_1 = &H53
Const CBWSignature_2 = &H42
Const CBWSignature_3 = &H43


Const CBWTag_0 = &H1
Const CBWTag_1 = &H2
Const CBWTag_2 = &H3
Const CBWTag_3 = &H4


'/////////////////// CBW signature

CBW(0) = CBWSignature_0
CBW(1) = CBWSignature_1
CBW(2) = CBWSignature_2
CBW(3) = CBWSignature_3

'/////////////////  CBW Tag

CBW(4) = CBWTag_0
CBW(5) = CBWTag_1
CBW(6) = CBWTag_2
CBW(7) = CBWTag_3

CBWDataTransferLen(0) = (CBWDataTransferLength Mod 256)
tmpV(0) = Int(CBWDataTransferLength / 256)
CBWDataTransferLen(1) = (tmpV(0) Mod 256)
tmpV(1) = Int(tmpV(0) / 256)
CBWDataTransferLen(2) = (tmpV(1) Mod 256)
tmpV(2) = Int((tmpV(1) / 256))
CBWDataTransferLen(3) = (tmpV(2) Mod 256)

CBW(8) = CBWDataTransferLen(0)  '00
CBW(9) = CBWDataTransferLen(1)  '08
CBW(10) = CBWDataTransferLen(2) '00
CBW(11) = CBWDataTransferLen(3) '00

'///////////////  CBW Flag
CBW(12) = &H80                 '80

'////////////// LUN
CBW(13) = Lun                    '00

'///////////// CBD Len
CBW(14) = &HA                '0a

'////////////  UFI command

CBW(15) = &H28
CBW(16) = Lun * 32
LBAByte(0) = (LBA Mod 256)
tmpV(0) = Int(LBA / 256)
LBAByte(1) = (tmpV(0) Mod 256)
tmpV(1) = Int(tmpV(0) / 256)
LBAByte(2) = (tmpV(1) Mod 256)
tmpV(2) = Int((tmpV(1) / 256))
LBAByte(3) = (tmpV(2) Mod 256)

CBW(17) = LBAByte(3)         '00
CBW(18) = LBAByte(2)         '00
CBW(19) = LBAByte(1)         '00
CBW(20) = LBAByte(0)         '40

'/////////////  Reverve
CBW(21) = 0

'//////////// Transfer Len

TransferLen = Int(CBWDataTransferLength / 512)

TransferLenLSB = (TransferLen Mod 256)
tmpV(0) = Int(TransferLen / 256)
TransferLenMSB = (tmpV(0) / 256)

CBW(22) = TransferLenMSB      '00
CBW(23) = TransferLenLSB      '04

For i = 24 To 30
    CBW(i) = 0
Next

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
 
Dim result As Long

'1. CBW command

 
result = WriteFile _
       (WriteHandle, _
       CBW(0), _
       31, _
       NumberOfBytesWritten, _
       0)    'out

 

'2. Readdata stage
 
result = ReadFile _
         (ReadHandle, _
          ReadData(0), _
         CBWDataTransferLength, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in

 
 

'3. CSW data
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 
 
'4. CSW status

If CSW(12) = 1 Then
    Read_DataAU9331 = 1
Else
     Read_DataAU9331 = 0
   
End If

 
End Function
Function OpenPipe9369() As Byte
Dim WritePathName As String
Dim ReadPathName As String

OpenPipe9369 = 0
'WritePathName = Left(DevicePathName, Len(DevicePathName) - 2) & "\PIPE0"   '
WritePathName = DevicePathName
'Debug.Print Lba
'Debug.Print "WritePathName="; WritePathName
WriteHandle9369 = CreateFile _
             (WritePathName, _
            GENERIC_READ Or GENERIC_WRITE, _
            (FILE_SHARE_READ Or FILE_SHARE_WRITE), _
             Security, _
             OPEN_EXISTING, _
             0&, _
            0)
'Debug.Print "write handle"; WriteHandle
If WriteHandle9369 = 0 Then
  OpenPipe9369 = 0
  Exit Function
End If
OpenPipe9369 = 1
End Function
Function ClosePipe9369() As Integer
On Error Resume Next
 
CloseHandle (WriteHandle9369)


End Function

Sub AU6254TestOld1(rv0 As Byte, rv1 As Byte, rv2 As Byte, rv3 As Byte, rv4 As Byte)
 On Error Resume Next
 Dim TestCounter As Integer
 Dim TimeInterval
 Dim HubEmuCounter As Integer
 Dim ContinueRun As Integer
 Dim PwrTime As Single
 
 ContinueRun = 1
 PwrTime = 0.7
 
 
 txtmsg.Text = ""
              
                If PCI7248InitFinish = 0 Then
                  PCI7248ExistAU6254
                End If
                          
               '====================================
               '            Hub exist test3
               ' ====================================
               '****************************
                'For HubEmuCounter = 1 To 5
               '****************************
                 
                  CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                  Print "1"
                  Call MsecDelay(0.5)
                  Cls
                  Print "2"
                  CardResult = DO_WritePort(card, Channel_P1A, &H1E)
                  Call MsecDelay(0.5) ' Hub Power on
                  If CardResult <> 0 Then
                      MsgBox "Power on fail"
                      End
                  End If
                  ReaderExist = 0
                  rv0 = 0
                  OldTimer = Timer
                  Cls
                  Print "3"
                  Do
                      Call MsecDelay(0.2)
                      DoEvents
                      rv0 = AU6254_GetDevice(0, 1, "6254")
                      TimeInterval = Timer - OldTimer
                  Loop While rv0 = 0 And TimeInterval < 5
                 
                  Print "rv0 ="; rv0; "TimeInterval="; TimeInterval
                  
                 '****************************
              '  Next HubEmuCounter
              '  MsgBox "Ok"
              '  Exit Sub
                '****************************
                
                If rv0 <> 1 Then 'Hub unknow
                     rv1 = 4
                     rv2 = 4
                     rv3 = 4
                     rv4 = 4
                     If ContinueRun = 0 Then
                        Exit Sub
                     End If
                End If
        
                Print "rv0="; rv0; "--- Hub Exist"
    
                '==========================  Hub Detect
                '1. usb 2.0 reader
                '2. usb 1.0 flash
                '3, usb 6610
                '4. usb keyboard
        
               '======== PWR ctrl
                CardResult = DO_WritePort(card, Channel_P1A, &H1C)
                Call MsecDelay(PwrTime)
                CardResult = DO_WritePort(card, Channel_P1A, &H18)
                Call MsecDelay(PwrTime)
                CardResult = DO_WritePort(card, Channel_P1A, &H10)
                Call MsecDelay(PwrTime)
                CardResult = DO_WritePort(card, Channel_P1A, &H0)
                Call MsecDelay(PwrTime)
                
                '===== 6335 test usb 2.0 speed
                HubPort = 0
                ReaderExist = 0
                rv1 = 0
                ClosePipe
                    'rv1 = AU6610Test
                rv1 = CBWTest_New_AU9254(0, 1, "6335")
                ClosePipe
 
                 
               Print "rv1="; rv1; "--- 2.0 speed and isochrous pipe test"
              
                '****************************
              '  Next HubEmuCounter
              '  MsgBox "Ok"
              '  Exit Sub
                '****************************
                
               If rv1 <> 1 Then
                   
                    rv2 = 4
                    rv3 = 4
                    rv4 = 4
                    If ContinueRun = 0 Then
                        Exit Sub
                    End If
                End If
                
                 Print "rv1="; rv1; "--- 2.0 speed and isochrous pipe test"
                 
                 
               '=================================
               '   test usb 1.1 flash
               '================================
            
                LBA = LBA + 1
                ReaderExist = 0
                ClosePipe6331
                  rv2 = CBWTest_New6331(0, 1, "6331")
                ClosePipe6331
           
                Print "rv2="; rv2; "--- 2.0 speed and Bulk r/w"
                Print " UsbSpeedTestResult="; UsbSpeedTestResult; "---usb speed error r/w"
                
                
                If rv2 <> 1 Then
                
                   If rv2 = 0 Then
                   rv2 = 2
                   End If
             
                   rv3 = 4
                   rv4 = 4
                   If ContinueRun = 0 Then
                        Exit Sub
                    End If
                End If
                 
                
               '=================================
               '   test 6610 for isochrous pipe
               '================================
               
               
                CardResult = DO_WritePort(card, Channel_P1A, &H0)     ' usb 2.0 falsh
               
                Call MsecDelay(1)
               
                HubPort = 1
                ReaderExist = 0
                ClosePipe
                rv3 = CBWTest_New_AU9254(0, 1, "9369")
                ClosePipe
             
             
                Print "rv3="; rv3; "--- 1.1 speed and isochrorous r/w"
               
               
                 CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                 If rv3 <> 1 Then
                        If rv3 = 0 Then
                          rv3 = 2
                        End If
                    
                       
                      rv4 = 4
                    If ContinueRun = 0 Then
                        Exit Sub
                    End If
                     
                  Else
                  
                     ' If LightON <> 243 Then
                     '     rv3 = 2
                      '   End If
                  
                     
                  End If
                
                
                 
                
                
                
               '=================================
               '   test usb 1.1 key board
               '================================
              
                 txtmsg.Text = ""
                 txtmsg.SetFocus
                 
                 HubPort = 2
                 ReaderExist = 0
                 ClosePipe
                   rv4 = CBWTest_New_AU9254(0, 1, "9462")
                 ClosePipe
                
                ' ReaderExist = 0
                ' ClosePipe
                '   rv4 = CBWTest_New_AU9254(0, 1, "9462")
                ' ClosePipe
        
                  '  CardResult = DO_WritePort(card, Channel_P1A, &HE)    ' usb 2.0 falsh
                
              '  Call MsecDelay(4)
                
                ' keybaord control
                 '  CardResult = DO_ReadPort(card, Channel_P1B, LightON)
                 '  Call MsecDelay(0.1)
                '   CardResult = DO_WritePort(card, Channel_P1CH, &H0)
                 '   Call MsecDelay(0.3)
                 
                 ' CardResult = DO_WritePort(card, Channel_P1CH, &H1)
                  '    Call MsecDelay(1.2)
                
                 
                    
                 ' CardResult = DO_WritePort(card, Channel_P1CH, &H0)
                 '   Call MsecDelay(0.2)
                  'CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' close power
                   
                 ' Print " LightON="; LightON
                  
                  
                   
                   
                   
               '    If InStr(txtmsg.Text, "..") <> 0 Then
                  
                '      rv4 = 1
                    
                '      If LightON = 238 Then
                '      rv4 = 1
                     
                 '    Else
                 '     Print "GPO Fail"
                     
                  '   End If
                     
               '  Else
               '    Print "Keyboard fail"
                    
                 
               '  End If
               
                  
                  If rv4 <> 1 Then
                      rv4 = 2
                   End If
                   Print "rv4="; rv4; "--- 12 MHZ speed and interrupt  r/w"
                
                
                    
                 If rv0 = 1 Or rv1 = 1 Or rv2 = 1 Or rv3 = 1 Or rv4 = 1 Then
                  
                    If rv0 <> 1 Then
                     ReaderExist = 0
                     rv0 = AU6254_GetDevice(0, 1, "6254")
                    End If
                    
                  If rv1 <> 1 Then
                  
                    HubPort = 0
                    ReaderExist = 0
                    ClosePipe
                    'rv1 = AU6610Test
                    rv1 = CBWTest_New_AU9254(0, 1, "6335")
                    ClosePipe
                    
                  End If
                  
                  
                  If rv2 <> 1 Then
                     ReaderExist = 0
                    ClosePipe6331
                    rv2 = CBWTest_New6331(0, 1, "6331")
                    ClosePipe6331
                  End If
                  
                  If rv3 <> 1 Then
                    HubPort = 1
                    ReaderExist = 0
                    ClosePipe
                    rv3 = CBWTest_New_AU9254(0, 1, "9369")
                    ClosePipe
                  End If
                  
                  If rv4 <> 1 Then
                  
                     HubPort = 2
                     ReaderExist = 0
                     ClosePipe
                     rv4 = CBWTest_New_AU9254(0, 1, "9462")
                     ClosePipe
        
                  End If
                 
                Print "RTrv0="; rv0
                Print "RTrv1="; rv1
                Print "RTrv2="; rv2
                Print "RTrv3="; rv3
                Print "RTrv4="; rv4

                
             End If
             
             
            '========== binning
             If rv0 = 0 Then   ' hub unknow  ,bin2
               rv1 = 4
               rv2 = 4
               rv3 = 4
               rv4 = 4
               Exit Sub
          End If
                
                
               If rv0 = 1 Then   ' hub unknow  , bin3
                 If rv1 = 0 Then  ' hub speed
                   rv1 = 2
                   rv2 = 4
                   rv3 = 4
                   rv4 = 4
                End If
                
                If rv1 * rv2 * rv3 * rv4 = 0 Then  ' bin4 down stream port unknow device
                    
                     If rv1 = 0 Then
                       PortFail = "port1 unknow device"
                    End If
                    
                       If rv2 = 0 Then
                       PortFail = "port2 unknow device"
                    End If
                    
                       If rv3 = 0 Then
                       PortFail = "port3 unknow device"
                    End If
                    
                       If rv4 = 0 Then
                       PortFail = "port4 unknow device"
                    End If
                      
                      
                    
                    rv1 = 1
                    rv2 = 1
                    rv3 = 2
                    rv4 = 4
                    
                  
                    
                    Exit Sub
                End If
                
                
                If rv2 = 2 Or rv4 = 2 Then ' bin5 down stream port unknow device
                    rv1 = 1
                    rv2 = 1
                    rv3 = 1
                    rv4 = 2
                End If
                
                Exit Sub
            End If
            
                


 End Sub

Function CBWTest_New9369(Lun As Byte, PreSlotStatus As Byte, Vid_PID As String) As Byte
Dim i As Integer
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long

   CBWDataTransferLength = 1024
 
'   For i = 0 To CBWDataTransferLength - 1
    
'         ReadData(i) = 0

'   Next

    If PreSlotStatus <> 1 Then
        CBWTest_New9369 = 4
        Exit Function
    End If
    '========================================
   
    CBWTest_New9369 = 0
    If LBA > 25 * 1024 Then
        LBA = 0
    End If
    '========================================
     TmpString = ""
    If ReaderExist = 0 Then
        Do
            DoEvents
            Call MsecDelay(0.1)
            TimerCounter = TimerCounter + 1
            TmpString = GetDeviceNameMulti9369(Vid_PID)
        Loop While TmpString = "" And TimerCounter < 10
    End If
    '=======================================
    If ReaderExist = 0 And TmpString <> "" Then
      ReaderExist = 1
    End If
    '=======================================
    If ReaderExist = 0 And TmpString = "" Then
      CBWTest_New9369 = 0   ' no readerExist
      ReaderExist = 0
      Exit Function
    End If
    '=======================================
    If OpenPipe9369 = 0 Then
      CBWTest_New9369 = 2   ' Write fail
      Exit Function
    End If
    Print DevicePathName
    '======================================
    
     ' for unitSpeed
    Dim ret As Integer
    
    ret = fnScsi2usb_TestUnitReady(WriteHandle9369)
    ret = fnScsi2usb_TestUnitReady(WriteHandle9369)
    ret = fnScsi2usb_Write(WriteHandle9369, 1, LBA, Pattern(0))
    ret = fnScsi2usb_Read(WriteHandle9369, 1, LBA)
    
    Receive_DataBuffer
    
    Dim i2 As Integer
    
    For i2 = 1 To 500
        If Pattern(i2) <> ReceiveDataBuffer(i2 + 1) Then
        CBWTest_New9369 = 2
        Exit Function
        End If
    Next i2
    
     CBWTest_New9369 = 1
   
  End Function
Sub AU9331Test()
 On Error Resume Next
 
             If PCI7248InitFinish = 0 Then
                  PCI7248Exist
              End If
              
            ' result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
            ' CardResult = DO_WritePort(card, Channel_P1B, &H0)   ' 0111 1111
            
            CardResult = DO_WritePort(card, Channel_P1A, &H1)   ' 0111 1111
           
                 
             Call MsecDelay(0.2)
            ' CardResult = DO_WritePort(card, Channel_P1B, &H2)   ' 0111 1111
             CardResult = DO_WritePort(card, Channel_P1A, &H2)   ' 0111 1111
                  
                 
             Call MsecDelay(1.2)
             
            
              LBA = LBA + 1
                
                
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
                
                
                Print LBA
                
                ClosePipe
                 rv0 = CBWTest_New_no_card(0, 1, "vid_058f")
                ClosePipe
                Call LabelMenu(0, rv0, 1)
                
              
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv0 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                ElseIf rv1 = WRITE_FAIL Then
                    CFWriteFail = CFWriteFail + 1
                    TestResult = "CF_WF"
                ElseIf rv1 = READ_FAIL Then
                    CFReadFail = CFReadFail + 1
                    TestResult = "CF_RF"
                ElseIf rv2 = WRITE_FAIL Then
                    XDWriteFail = XDWriteFail + 1
                    TestResult = "XD_WF"
                ElseIf rv2 = READ_FAIL Then
                    XDReadFail = XDReadFail + 1
                    TestResult = "XD_RF"
                 ElseIf rv3 = WRITE_FAIL Then
                    MSWriteFail = MSWriteFail + 1
                    TestResult = "MS_WF"
                ElseIf rv3 = READ_FAIL Then
                    MSReadFail = MSReadFail + 1
                    TestResult = "MS_RF"
                ElseIf rv0 = PASS Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
                
                Print "Test Result"; TestResult
                       
       
                 
                Print rv0, " \\SD :0 Unknow device, 1 pass ,2 card change bit fail"
            
                 
                If TestResult = "PASS" Then
                  
                '   result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                '  CardResult = DO_WritePort(card, Channel_P1B, &H1)   ' 0111 1111
                  
                   CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' 0110 0100
                    
                   Call MsecDelay(0.5)
                 
                 '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                 '
                 '  R/W test
                 '
                 '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                 
                
                'initial return value
                
                
                 
                rv1 = 0
                rv2 = 0
                rv3 = 0
               
                Label3.BackColor = RGB(255, 255, 255)
                Label4.BackColor = RGB(255, 255, 255)
                Label5.BackColor = RGB(255, 255, 255)
                Label6.BackColor = RGB(255, 255, 255)
                Label7.BackColor = RGB(255, 255, 255)
                Label8.BackColor = RGB(255, 255, 255)
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                 rv1 = CBWTest_NewAU9331(0, 1, "vid_058f")
               
                ClosePipe
                Call LabelMenu(0, rv1, rv0)
                 ' CardResult = DO_ReadPort(card, Channel_P1B, LightOFF)
                '                 If LightOFF <> 192 Then
                '                   UsbSpeedTestResult = GPO_FAIL
                '                  rv1 = 2
                '                 End If
                '   Call LabelMenu(0, rv1, rv0)
                      
     
                
                Print rv1, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print "LBA="; LBA
                
              
                        
                        If rv0 = UNKNOW Then
                           UnknowDeviceFail = UnknowDeviceFail + 1
                           TestResult = "UNKNOW"
                        ElseIf rv0 = WRITE_FAIL Then
                            SDWriteFail = SDWriteFail + 1
                            TestResult = "SD_WF"
                        ElseIf rv0 = READ_FAIL Then
                            SDReadFail = SDReadFail + 1
                            TestResult = "SD_RF"
                        ElseIf rv1 = WRITE_FAIL Then
                            CFWriteFail = CFWriteFail + 1
                            TestResult = "CF_WF"
                        ElseIf rv1 = READ_FAIL Then
                            CFReadFail = CFReadFail + 1
                            TestResult = "CF_RF"
                        ElseIf rv2 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                        ElseIf rv2 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                         ElseIf rv3 = WRITE_FAIL Or rv4 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv3 = READ_FAIL Or rv4 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                        ElseIf rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
                
               
                  
                End If
                CardResult = DO_WritePort(card, Channel_P1A, &HFF)
     
   


 End Sub


 Sub AU6254TestAlcorAU6256BLF20(rv0 As Byte, rv1 As Byte, rv2 As Byte, rv3 As Byte, rv4 As Byte)
 On Error Resume Next
 Dim TestCounter As Integer
 Dim TimeInterval
 Dim HubEmuCounter As Integer
 Dim ContinueRun As Integer
 Dim PwrTime As Single
 Dim Au6254Speed As Integer
 Dim ResetBit As Integer
 Dim FailPosition As Integer
 
 ContinueRun = 1
 PwrTime = 0.5
 ResetBit = 128
 
 txtmsg.Text = ""
              
                If PCI7248InitFinish = 0 Then
                  PCI7248ExistAU6254
                End If
                          
               '====================================
               '            Hub exist test3
               ' ====================================
               '****************************
                'For HubEmuCounter = 1 To 5
               '****************************
                    WinExec "off.exe", 0
                    CardResult = DO_WritePort(card, Channel_P1CL, &H4)
                '  CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                 ' Print "1"
                  Call MsecDelay(0.3) ' system let hub driver unload time
               
                  '============= dETECT FOR Disconnect fail
       
                  
                  Cls
                  Print "2"
                    WinExec "on.exe", 0
                  CardResult = DO_WritePort(card, Channel_P1A, &H1E + ResetBit)
                  Call MsecDelay(CSng(Text25.Text)) ' Hub Power on
                  If CardResult <> 0 Then
                      MsgBox "Power on fail"
                      End
                  End If
                  ReaderExist = 0
                  rv0 = 0
                  OldTimer = Timer
                  Cls
                  Print "3"
                  Do
                      Call MsecDelay(0.2)
                      DoEvents
                       rv0 = AU6254_GetDevice(0, 1, "6254")
                       
                      TimeInterval = Timer - OldTimer
                  Loop While rv0 = 0 And TimeInterval < 5
                 
                  Print "rv0 ="; rv0; "TimeInterval="; TimeInterval
                  
                 '****************************
              '  Next HubEmuCounter
              '  MsgBox "Ok"
              '  Exit Sub
                '****************************
                
                If rv0 <> 1 Then 'Hub unknow
                     rv1 = 4
                     rv2 = 4
                     rv3 = 4
                     rv4 = 4
                     If ContinueRun = 0 Then
                        Exit Sub
                     End If
                End If
        
                Print "rv0="; rv0; "--- Hub Exist"
    
                '==========================  Hub Detect
                '1. usb 2.0 reader
                '2. usb 1.0 flash
                '3, usb 6610
                '4. usb keyboard
        
               '======== PWR ctrl
               If ChipName <> "AU6254XLT20" Then
                CardResult = DO_WritePort(card, Channel_P1A, &H1C + ResetBit)
               End If
                Call MsecDelay(PwrTime)
                
                '===============   must at here to test speed , otherwise driver will overlap the test result
                
                HubPort = 0
                ReaderExist = 0
                rv1 = 0
                ClosePipe
                'rv1 = AU6610Test
                rv1 = CBWTest_New_AU9254(0, 1, "6377")
                ClosePipe
                Au6254Speed = UsbSpeedTestResult
                 
                Print "rv1="; rv1; "--- 2.0 speed test"
                
                
                   
               If rv1 <> 1 Then
                   
                    rv2 = 4
                    rv3 = 4
                    rv4 = 4
                    
                    
                        If rv1 >= 2 Then  ' speed error
                             Exit Sub
                        End If
                    
                    If ContinueRun = 0 Then
                        Exit Sub
                    End If
                End If
                
                 Print "rv1="; rv1; "--- 2.0 speed and isochrous pipe test"
               
                If ChipName <> "AU6254XLT20" Then
                CardResult = DO_WritePort(card, Channel_P1A, &H18 + ResetBit)
               
                Call MsecDelay(PwrTime)
                CardResult = DO_WritePort(card, Channel_P1A, &H10 + ResetBit)
                Call MsecDelay(PwrTime)
                CardResult = DO_WritePort(card, Channel_P1A, &H0 + ResetBit)
                Call MsecDelay(PwrTime)
                 
                Else
                
                CardResult = DO_WritePort(card, Channel_P1A, &H16 + ResetBit)
                Call MsecDelay(PwrTime)
                
                End If
                
                
                
                 
                '===== 6335 test usb 2.0 speed
              
                   
                 
                
                '****************************
              '  Next HubEmuCounter
              '  MsgBox "Ok"
              '  Exit Sub
                '****************************
             
                 
                 
               '=================================
               '   test usb 1.1 flash
               '================================
                HubPort = 1
                LBA = LBA + 1
                ReaderExist = 0
                ClosePipe
                  rv2 = CBWTest_New_AU9254(0, 1, "6377")
                ClosePipe
           
                Print "rv2="; rv2; "--- 2.0 speed and Bulk r/w"
                Print " UsbSpeedTestResult="; UsbSpeedTestResult; "---usb speed error r/w"
                
                
                If rv2 <> 1 Then
    
                   rv3 = 4
                   rv4 = 4
                   If ContinueRun = 0 Then
                        Exit Sub
                    End If
                End If
                 
                
               '=================================
               '   test 6610 for isochrous pipe
               '================================
               
               If ChipName = "AU6254XLT20" Then
               HubPort = 0
               Else
                HubPort = 2
               End If
               
                ReaderExist = 0
                ClosePipe
                rv3 = CBWTest_New_AU9254(0, 1, "9360")
                rv3 = CBWTest_New_8_Sector(0, 1, FailPosition, 20)
                ClosePipe
         
                Print "rv3="; rv3; "--- 1.1 speed and isochrorous r/w"
               
               
                 CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                 If rv3 <> 1 Then
                  rv4 = 4
                    If ContinueRun = 0 Then
                        Exit Sub
                    End If
                     
                  Else
                  
                     ' If LightON <> 243 Then
                     '     rv3 = 2
                      '   End If
                  
                     
                  End If
                
                
                 
                
                
                
               '=================================
               '   test usb 1.1 key board
               '================================
                  
                 txtmsg.Text = ""
                 txtmsg.SetFocus
                 
                 HubPort = 3
                 ReaderExist = 0
                 ClosePipe
                   rv4 = CBWTest_New_AU9254(0, 1, "413c")
                 ClosePipe
                
                
                   Print "rv4="; rv4; "--- 12 MHZ speed and interrupt  r/w"
                
                
                    
                 If rv0 = 1 Or rv1 = 1 Or rv2 = 1 Or rv3 = 1 Or rv4 = 1 Then
                 
                 
                    If rv0 * rv1 * rv2 * rv3 * rv4 = 1 Then
                      Rv1ContinueFail = 0
                    End If
                    
                  
                    If rv0 <> 1 Then
                    'Rv1ContinueFail = 0
                     ReaderExist = 0
                     rv0 = AU6254_GetDevice(0, 1, "6254")
                    End If
                    
                  If rv1 <> 1 Then
    
                    HubPort = 0
                    ReaderExist = 0
                    ClosePipe
                    'rv1 = AU6610Test
                    rv1 = CBWTest_New_AU9254(0, 1, "6377")
                    ClosePipe
                     Au6254Speed = UsbSpeedTestResult
     
                     
                     
                  End If
                  
                  
                  If rv2 <> 1 Then
                    ' Rv1ContinueFail = 0
                    
                     HubPort = 1
                     ReaderExist = 0
                    ClosePipe
                    rv2 = CBWTest_New_AU9254(0, 1, "6377")
                    ClosePipe
                  End If
                  
                  If rv3 <> 1 Then
                    'Rv1ContinueFail = 0
                    HubPort = 2
                    ReaderExist = 0
                      ClosePipe
                     rv3 = CBWTest_New_AU9254(0, 1, "9360")
                     rv3 = CBWTest_New_8_Sector(0, 1, FailPosition, 20)
                    ClosePipe
              
                  End If
                  
                  If rv4 <> 1 Then
                   '  Rv1ContinueFail = 0
                     HubPort = 3
                     ReaderExist = 0
                     ClosePipe
                     rv4 = CBWTest_New_AU9254(0, 1, "413c")
                     ClosePipe
        
                  End If
                  
                  
                    If rv1 + rv2 + rv3 + rv4 = 0 Then
                     
                      Rv1ContinueFail = Rv1ContinueFail + 1
                     End If
                      
                     
                     
                    If Rv1ContinueFail > 1 Then  ' use another AU6335 plug into to solve the hub hang probelm
                    CardResult = DO_WritePort(card, Channel_P1CL, &H0)
                    Call MsecDelay(2.5)
                    CardResult = DO_WritePort(card, Channel_P1CL, &H4)
                    Call MsecDelay(2)
                   End If
                 
                Print "RTrv0="; rv0
                Print "RTrv1="; rv1
                Print "RTrv2="; rv2
                Print "RTrv3="; rv3
                Print "RTrv4="; rv4

                
             End If
             
             
            '========== binning
             If rv0 = 0 Then   ' hub unknow  ,bin2
               rv1 = 4
               rv2 = 4
               rv3 = 4
               rv4 = 4
               AU6254TestMsg = "Hub Unknow"
               Exit Sub
          End If
          
        
               If rv0 = 1 Then   ' hub unknow  , bin3
                 If Au6254Speed = 2 Then     ' hub speed
                   rv1 = 2
                   rv2 = 4
                   rv3 = 4
                   rv4 = 4
                    AU6254TestMsg = "USB 2.0 speed error"
                   Exit Sub
                End If
                
                If rv1 * rv2 * rv3 * rv4 = 0 Then  ' bin4 down stream port unknow device
                     PortFail = ""
                     If rv1 = 0 Then
                       PortFail = PortFail & "port1 unknow,"
                    End If
                    
                       If rv2 = 0 Then
                       PortFail = PortFail & "port2 unknow,"
                    End If
                    
                       If rv3 = 0 Then
                       PortFail = PortFail & "port3 unknow,"
                    End If
                    
                       If rv4 = 0 Then
                       PortFail = PortFail & "port4 unknow,"
                    End If
                      
                      
                    
                    rv1 = 1
                    rv2 = 1
                    rv3 = 2
                    rv4 = 4
                    
                    AU6254TestMsg = PortFail
                    
                    Exit Sub
                End If
                
                
                
                
                If rv2 >= 2 Then   ' bin5 down stream port unknow device
                    rv1 = 1
                    rv2 = 1
                    rv3 = 1
                    rv4 = 2
                    AU6254TestMsg = "2.0 Reader SD R/W fail"
                    Exit Sub
                End If
                
                  
                If rv3 >= 2 Then   ' bin5 down stream port unknow device
                    rv1 = 1
                    rv2 = 1
                    rv3 = 1
                    rv4 = 2
                    AU6254TestMsg = "1.0 Reader SD R/W fail"
                    Exit Sub
                End If
                
                
                CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                
                 
             '  If LightON <> 225 And LightON <> 224 Then ' bin5 : light on fail
             '      rv1 = 1
             '      rv2 = 1
             '      rv3 = 1
             '      rv4 = 2
             '      AU6254TestMsg = "GPO R/W fail"
             '  End If
                Exit Sub
            End If
            
                


 End Sub
 
 Sub AU6254TestAlcor(rv0 As Byte, rv1 As Byte, rv2 As Byte, rv3 As Byte, rv4 As Byte)
 On Error Resume Next
 Dim TestCounter As Integer
 Dim TimeInterval
 Dim HubEmuCounter As Integer
 Dim ContinueRun As Integer
 Dim PwrTime As Single
 Dim Au6254Speed As Integer
 Dim ResetBit As Integer
 Dim FailPosition As Integer
 
 ContinueRun = 1
 PwrTime = 0.5
 ResetBit = 128
 
 txtmsg.Text = ""
              
                If PCI7248InitFinish = 0 Then
                  PCI7248ExistAU6254
                End If
                          
               '====================================
               '            Hub exist test3
               ' ====================================
               '****************************
                'For HubEmuCounter = 1 To 5
               '****************************
                    WinExec "off.exe", 0
                    CardResult = DO_WritePort(card, Channel_P1CL, &H4)
                '  CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                 ' Print "1"
                  Call MsecDelay(0.3) ' system let hub driver unload time
               
                  '============= dETECT FOR Disconnect fail
       
                  
                  Cls
                  Print "2"
                    WinExec "on.exe", 0
                  CardResult = DO_WritePort(card, Channel_P1A, &H1E + ResetBit)
                  Call MsecDelay(CSng(Text25.Text)) ' Hub Power on
                  If CardResult <> 0 Then
                      MsgBox "Power on fail"
                      End
                  End If
                  ReaderExist = 0
                  rv0 = 0
                  OldTimer = Timer
                  Cls
                  Print "3"
                  Do
                      Call MsecDelay(0.2)
                      DoEvents
                       rv0 = AU6254_GetDevice(0, 1, "6254")
                       
                      TimeInterval = Timer - OldTimer
                  Loop While rv0 = 0 And TimeInterval < 5
                 
                  Print "rv0 ="; rv0; "TimeInterval="; TimeInterval
                  
                 '****************************
              '  Next HubEmuCounter
              '  MsgBox "Ok"
              '  Exit Sub
                '****************************
                
                If rv0 <> 1 Then 'Hub unknow
                     rv1 = 4
                     rv2 = 4
                     rv3 = 4
                     rv4 = 4
                     If ContinueRun = 0 Then
                        Exit Sub
                     End If
                End If
        
                Print "rv0="; rv0; "--- Hub Exist"
    
                '==========================  Hub Detect
                '1. usb 2.0 reader
                '2. usb 1.0 flash
                '3, usb 6610
                '4. usb keyboard
        
               '======== PWR ctrl
               If ChipName <> "AU6254XLT20" Then
                CardResult = DO_WritePort(card, Channel_P1A, &H1C + ResetBit)
               End If
                Call MsecDelay(PwrTime)
                
                '===============   must at here to test speed , otherwise driver will overlap the test result
                
                HubPort = 0
                ReaderExist = 0
                rv1 = 0
                ClosePipe
                'rv1 = AU6610Test
                rv1 = CBWTest_New_AU9254(0, 1, "6377")
                ClosePipe
                Au6254Speed = UsbSpeedTestResult
                 
                Print "rv1="; rv1; "--- 2.0 speed test"
                
                
                   
               If rv1 <> 1 Then
                   
                    rv2 = 4
                    rv3 = 4
                    rv4 = 4
                    
                    
                        If rv1 >= 2 Then  ' speed error
                             Exit Sub
                        End If
                    
                    If ContinueRun = 0 Then
                        Exit Sub
                    End If
                End If
                
                 Print "rv1="; rv1; "--- 2.0 speed and isochrous pipe test"
               
                If ChipName <> "AU6254XLT20" Then
                CardResult = DO_WritePort(card, Channel_P1A, &H18 + ResetBit)
               
                Call MsecDelay(PwrTime)
                CardResult = DO_WritePort(card, Channel_P1A, &H10 + ResetBit)
                Call MsecDelay(PwrTime)
                CardResult = DO_WritePort(card, Channel_P1A, &H0 + ResetBit)
                Call MsecDelay(PwrTime)
                 
                Else
                
                CardResult = DO_WritePort(card, Channel_P1A, &H16 + ResetBit)
                Call MsecDelay(PwrTime)
                
                End If
                
                
                
                 
                '===== 6335 test usb 2.0 speed
              
                   
                 
                
                '****************************
              '  Next HubEmuCounter
              '  MsgBox "Ok"
              '  Exit Sub
                '****************************
             
                 
                 
               '=================================
               '   test usb 1.1 flash
               '================================
                HubPort = 1
                LBA = LBA + 1
                ReaderExist = 0
                ClosePipe
                  rv2 = CBWTest_New_AU9254(0, 1, "6377")
                ClosePipe
           
                Print "rv2="; rv2; "--- 2.0 speed and Bulk r/w"
                Print " UsbSpeedTestResult="; UsbSpeedTestResult; "---usb speed error r/w"
                
                
                If rv2 <> 1 Then
    
                   rv3 = 4
                   rv4 = 4
                   If ContinueRun = 0 Then
                        Exit Sub
                    End If
                End If
                 
                
               '=================================
               '   test 6610 for isochrous pipe
               '================================
               
               If ChipName = "AU6254XLT20" Then
               HubPort = 0
               Else
                HubPort = 2
               End If
               
                ReaderExist = 0
                ClosePipe
                rv3 = CBWTest_New_AU9254(0, 1, "9360")
                rv3 = CBWTest_New_8_Sector(0, 1, FailPosition, 20)
                ClosePipe
         
                Print "rv3="; rv3; "--- 1.1 speed and isochrorous r/w"
               
               
                 CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                 If rv3 <> 1 Then
                  rv4 = 4
                    If ContinueRun = 0 Then
                        Exit Sub
                    End If
                     
                  Else
                  
                     ' If LightON <> 243 Then
                     '     rv3 = 2
                      '   End If
                  
                     
                  End If
                
                
                 
                
                
                
               '=================================
               '   test usb 1.1 key board
               '================================
                  
                 txtmsg.Text = ""
                 txtmsg.SetFocus
                 
                 HubPort = 3
                 ReaderExist = 0
                 ClosePipe
                   rv4 = CBWTest_New_AU9254(0, 1, "413c")
                 ClosePipe
                
                
                   Print "rv4="; rv4; "--- 12 MHZ speed and interrupt  r/w"
                
                
                    
                 If rv0 = 1 Or rv1 = 1 Or rv2 = 1 Or rv3 = 1 Or rv4 = 1 Then
                 
                 
                    If rv0 * rv1 * rv2 * rv3 * rv4 = 1 Then
                      Rv1ContinueFail = 0
                    End If
                    
                  
                    If rv0 <> 1 Then
                    'Rv1ContinueFail = 0
                     ReaderExist = 0
                     rv0 = AU6254_GetDevice(0, 1, "6254")
                    End If
                    
                  If rv1 <> 1 Then
    
                    HubPort = 0
                    ReaderExist = 0
                    ClosePipe
                    'rv1 = AU6610Test
                    rv1 = CBWTest_New_AU9254(0, 1, "6377")
                    ClosePipe
                     Au6254Speed = UsbSpeedTestResult
     
                     
                     
                  End If
                  
                  
                  If rv2 <> 1 Then
                    ' Rv1ContinueFail = 0
                    
                     HubPort = 1
                     ReaderExist = 0
                    ClosePipe
                    rv2 = CBWTest_New_AU9254(0, 1, "6377")
                    ClosePipe
                  End If
                  
                  If rv3 <> 1 Then
                    'Rv1ContinueFail = 0
                    HubPort = 2
                    ReaderExist = 0
                      ClosePipe
                     rv3 = CBWTest_New_AU9254(0, 1, "9360")
                     rv3 = CBWTest_New_8_Sector(0, 1, FailPosition, 20)
                    ClosePipe
              
                  End If
                  
                  If rv4 <> 1 Then
                   '  Rv1ContinueFail = 0
                     HubPort = 3
                     ReaderExist = 0
                     ClosePipe
                     rv4 = CBWTest_New_AU9254(0, 1, "413c")
                     ClosePipe
        
                  End If
                  
                  
                    If rv1 + rv2 + rv3 + rv4 = 0 Then
                     
                      Rv1ContinueFail = Rv1ContinueFail + 1
                     End If
                      
                     
                     
                    If Rv1ContinueFail > 1 Then  ' use another AU6335 plug into to solve the hub hang probelm
                    CardResult = DO_WritePort(card, Channel_P1CL, &H0)
                    Call MsecDelay(2.5)
                    CardResult = DO_WritePort(card, Channel_P1CL, &H4)
                    Call MsecDelay(2)
                   End If
                 
                Print "RTrv0="; rv0
                Print "RTrv1="; rv1
                Print "RTrv2="; rv2
                Print "RTrv3="; rv3
                Print "RTrv4="; rv4

                
             End If
             
             
            '========== binning
             If rv0 = 0 Then   ' hub unknow  ,bin2
               rv1 = 4
               rv2 = 4
               rv3 = 4
               rv4 = 4
               AU6254TestMsg = "Hub Unknow"
               Exit Sub
          End If
          
        
               If rv0 = 1 Then   ' hub unknow  , bin3
                 If Au6254Speed = 2 Then     ' hub speed
                   rv1 = 2
                   rv2 = 4
                   rv3 = 4
                   rv4 = 4
                    AU6254TestMsg = "USB 2.0 speed error"
                   Exit Sub
                End If
                
                If rv1 * rv2 * rv3 * rv4 = 0 Then  ' bin4 down stream port unknow device
                     PortFail = ""
                     If rv1 = 0 Then
                       PortFail = PortFail & "port1 unknow,"
                    End If
                    
                       If rv2 = 0 Then
                       PortFail = PortFail & "port2 unknow,"
                    End If
                    
                       If rv3 = 0 Then
                       PortFail = PortFail & "port3 unknow,"
                    End If
                    
                       If rv4 = 0 Then
                       PortFail = PortFail & "port4 unknow,"
                    End If
                      
                      
                    
                    rv1 = 1
                    rv2 = 1
                    rv3 = 2
                    rv4 = 4
                    
                    AU6254TestMsg = PortFail
                    
                    Exit Sub
                End If
                
                
                
                
                If rv2 >= 2 Then   ' bin5 down stream port unknow device
                    rv1 = 1
                    rv2 = 1
                    rv3 = 1
                    rv4 = 2
                    AU6254TestMsg = "2.0 Reader SD R/W fail"
                    Exit Sub
                End If
                
                  
                If rv3 >= 2 Then   ' bin5 down stream port unknow device
                    rv1 = 1
                    rv2 = 1
                    rv3 = 1
                    rv4 = 2
                    AU6254TestMsg = "1.0 Reader SD R/W fail"
                    Exit Sub
                End If
                
                
                CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                
                 
               If LightOn <> 225 And LightOn <> 224 Then ' bin5 : light on fail
                   rv1 = 1
                   rv2 = 1
                   rv3 = 1
                   rv4 = 2
                   AU6254TestMsg = "GPO R/W fail"
               End If
                Exit Sub
            End If
            
                


 End Sub
 Sub AU6254TestAlcorAU6254AS(rv0 As Byte, rv1 As Byte, rv2 As Byte, rv3 As Byte, rv4 As Byte)
 On Error Resume Next
 Dim TestCounter As Integer
 Dim TimeInterval
 Dim HubEmuCounter As Integer
 Dim ContinueRun As Integer
 Dim PwrTime As Single
 Dim Au6254Speed As Integer
 Dim ResetBit As Integer
 Dim FailPosition As Integer
 
 ContinueRun = 1
 PwrTime = 0.5
 ResetBit = 128
 
 txtmsg.Text = ""
              
                If PCI7248InitFinish = 0 Then
                  PCI7248ExistAU6254
                End If
                          
               '====================================
               '            Hub exist test3
               ' ====================================
               '****************************
                'For HubEmuCounter = 1 To 5
               '****************************
                    WinExec "off.exe", 0
                    CardResult = DO_WritePort(card, Channel_P1CL, &H4)
                '  CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                 ' Print "1"
                  Call MsecDelay(0.3) ' system let hub driver unload time
               
                  '============= dETECT FOR Disconnect fail
       
                  
                  Cls
                  Print "2"
                    WinExec "on.exe", 0
                  CardResult = DO_WritePort(card, Channel_P1A, &H1E + ResetBit)
                  Call MsecDelay(CSng(Text25.Text)) ' Hub Power on
                  If CardResult <> 0 Then
                      MsgBox "Power on fail"
                      End
                  End If
                  ReaderExist = 0
                  rv0 = 0
                  OldTimer = Timer
                  Cls
                  Print "3"
                  Do
                      Call MsecDelay(0.2)
                      DoEvents
                       rv0 = AU6254_GetDevice(0, 1, "6254")
                       
                      TimeInterval = Timer - OldTimer
                  Loop While rv0 = 0 And TimeInterval < 5
                 
                  Print "rv0 ="; rv0; "TimeInterval="; TimeInterval
                  
                 '****************************
              '  Next HubEmuCounter
              '  MsgBox "Ok"
              '  Exit Sub
                '****************************
                
                If rv0 <> 1 Then 'Hub unknow
                     rv1 = 4
                     rv2 = 4
                     rv3 = 4
                     rv4 = 4
                     If ContinueRun = 0 Then
                        Exit Sub
                     End If
                End If
        
                Print "rv0="; rv0; "--- Hub Exist"
    
                '==========================  Hub Detect
                '1. usb 2.0 reader
                '2. usb 1.0 flash
                '3, usb 6610
                '4. usb keyboard
        
               '======== PWR ctrl
               If ChipName <> "AU6254XLT20" Then
                CardResult = DO_WritePort(card, Channel_P1A, &H1C + ResetBit)
               End If
                Call MsecDelay(PwrTime)
                
                '===============   must at here to test speed , otherwise driver will overlap the test result
                
                HubPort = 0
                ReaderExist = 0
                rv1 = 0
                ClosePipe
                'rv1 = AU6610Test
                rv1 = CBWTest_New_AU9254(0, 1, "6377")
                ClosePipe
                Au6254Speed = UsbSpeedTestResult
                 
                Print "rv1="; rv1; "--- 2.0 speed test"
                
                
                   
               If rv1 <> 1 Then
                   
                    rv2 = 4
                    rv3 = 4
                    rv4 = 4
                    
                    
                        If rv1 >= 2 Then  ' speed error
                             Exit Sub
                        End If
                    
                    If ContinueRun = 0 Then
                        Exit Sub
                    End If
                End If
                
                 Print "rv1="; rv1; "--- 2.0 speed and isochrous pipe test"
               
                If ChipName <> "AU6254XLT20" Then
                CardResult = DO_WritePort(card, Channel_P1A, &H18 + ResetBit)
               
                Call MsecDelay(PwrTime)
                CardResult = DO_WritePort(card, Channel_P1A, &H10 + ResetBit)
                Call MsecDelay(PwrTime)
                CardResult = DO_WritePort(card, Channel_P1A, &H0 + ResetBit)
                Call MsecDelay(PwrTime)
                 
                Else
                
                CardResult = DO_WritePort(card, Channel_P1A, &H16 + ResetBit)
                Call MsecDelay(PwrTime)
                
                End If
                
                
                
                 
                '===== 6335 test usb 2.0 speed
              
                   
                 
                
                '****************************
              '  Next HubEmuCounter
              '  MsgBox "Ok"
              '  Exit Sub
                '****************************
             
                 
                 
               '=================================
               '   test usb 1.1 flash
               '================================
                HubPort = 1
                LBA = LBA + 1
                ReaderExist = 0
                ClosePipe
                  rv2 = CBWTest_New_AU9254(0, 1, "6377")
                ClosePipe
           
                Print "rv2="; rv2; "--- 2.0 speed and Bulk r/w"
                Print " UsbSpeedTestResult="; UsbSpeedTestResult; "---usb speed error r/w"
                
                
                If rv2 <> 1 Then
    
                   rv3 = 4
                   rv4 = 4
                   If ContinueRun = 0 Then
                        Exit Sub
                    End If
                End If
                 
                
               '=================================
               '   test 6610 for isochrous pipe
               '================================
               
               If ChipName = "AU6254XLT20" Then
               HubPort = 0
               Else
                HubPort = 2
               End If
               
                ReaderExist = 0
                ClosePipe
                rv3 = CBWTest_New_AU9254(0, 1, "9360")
                rv3 = CBWTest_New_8_Sector(0, 1, FailPosition, 20)
                ClosePipe
         
                Print "rv3="; rv3; "--- 1.1 speed and isochrorous r/w"
               
               
                 CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                 If rv3 <> 1 Then
                  rv4 = 4
                    If ContinueRun = 0 Then
                        Exit Sub
                    End If
                     
                  Else
                  
                     ' If LightON <> 243 Then
                     '     rv3 = 2
                      '   End If
                  
                     
                  End If
                
                
                 
                
                
                
               '=================================
               '   test usb 1.1 key board
               '================================
                  
                 txtmsg.Text = ""
                 txtmsg.SetFocus
                 
                 HubPort = 3
                 ReaderExist = 0
                 ClosePipe
                   rv4 = CBWTest_New_AU9254(0, 1, "413c")
                 ClosePipe
                
                
                   Print "rv4="; rv4; "--- 12 MHZ speed and interrupt  r/w"
                
                
                    
                 If rv0 = 1 Or rv1 = 1 Or rv2 = 1 Or rv3 = 1 Or rv4 = 1 Then
                 
                 
                    If rv0 * rv1 * rv2 * rv3 * rv4 = 1 Then
                      Rv1ContinueFail = 0
                    End If
                    
                  
                    If rv0 <> 1 Then
                    'Rv1ContinueFail = 0
                     ReaderExist = 0
                     rv0 = AU6254_GetDevice(0, 1, "6254")
                    End If
                    
                  If rv1 <> 1 Then
    
                    HubPort = 0
                    ReaderExist = 0
                    ClosePipe
                    'rv1 = AU6610Test
                    rv1 = CBWTest_New_AU9254(0, 1, "6377")
                    ClosePipe
                     Au6254Speed = UsbSpeedTestResult
     
                     
                     
                  End If
                  
                  
                  If rv2 <> 1 Then
                    ' Rv1ContinueFail = 0
                    
                     HubPort = 1
                     ReaderExist = 0
                    ClosePipe
                    rv2 = CBWTest_New_AU9254(0, 1, "6377")
                    ClosePipe
                  End If
                  
                  If rv3 <> 1 Then
                    'Rv1ContinueFail = 0
                    HubPort = 2
                    ReaderExist = 0
                      ClosePipe
                     rv3 = CBWTest_New_AU9254(0, 1, "9360")
                     rv3 = CBWTest_New_8_Sector(0, 1, FailPosition, 20)
                    ClosePipe
              
                  End If
                  
                  If rv4 <> 1 Then
                   '  Rv1ContinueFail = 0
                     HubPort = 3
                     ReaderExist = 0
                     ClosePipe
                     rv4 = CBWTest_New_AU9254(0, 1, "413c")
                     ClosePipe
        
                  End If
                  
                  
                    If rv1 + rv2 + rv3 + rv4 = 0 Then
                     
                      Rv1ContinueFail = Rv1ContinueFail + 1
                     End If
                      
                     
                     
                    If Rv1ContinueFail > 1 Then  ' use another AU6335 plug into to solve the hub hang probelm
                    CardResult = DO_WritePort(card, Channel_P1CL, &H0)
                    Call MsecDelay(2.5)
                    CardResult = DO_WritePort(card, Channel_P1CL, &H4)
                    Call MsecDelay(2)
                   End If
                 
                Print "RTrv0="; rv0
                Print "RTrv1="; rv1
                Print "RTrv2="; rv2
                Print "RTrv3="; rv3
                Print "RTrv4="; rv4

                
             End If
             
             
            '========== binning
             If rv0 = 0 Then   ' hub unknow  ,bin2
               rv1 = 4
               rv2 = 4
               rv3 = 4
               rv4 = 4
               AU6254TestMsg = "Hub Unknow"
               Exit Sub
          End If
          
        
               If rv0 = 1 Then   ' hub unknow  , bin3
                 If Au6254Speed = 2 Then     ' hub speed
                   rv1 = 2
                   rv2 = 4
                   rv3 = 4
                   rv4 = 4
                    AU6254TestMsg = "USB 2.0 speed error"
                   Exit Sub
                End If
                
                If rv1 * rv2 * rv3 * rv4 = 0 Then  ' bin4 down stream port unknow device
                     PortFail = ""
                     If rv1 = 0 Then
                       PortFail = PortFail & "port1 unknow,"
                    End If
                    
                       If rv2 = 0 Then
                       PortFail = PortFail & "port2 unknow,"
                    End If
                    
                       If rv3 = 0 Then
                       PortFail = PortFail & "port3 unknow,"
                    End If
                    
                       If rv4 = 0 Then
                       PortFail = PortFail & "port4 unknow,"
                    End If
                      
                      
                    
                    rv1 = 1
                    rv2 = 1
                    rv3 = 2
                    rv4 = 4
                    
                    AU6254TestMsg = PortFail
                    
                    Exit Sub
                End If
                
                
                
                
                If rv2 >= 2 Then   ' bin5 down stream port unknow device
                    rv1 = 1
                    rv2 = 1
                    rv3 = 1
                    rv4 = 2
                    AU6254TestMsg = "2.0 Reader SD R/W fail"
                    Exit Sub
                End If
                
                  
                If rv3 >= 2 Then   ' bin5 down stream port unknow device
                    rv1 = 1
                    rv2 = 1
                    rv3 = 1
                    rv4 = 2
                    AU6254TestMsg = "1.0 Reader SD R/W fail"
                    Exit Sub
                End If
                
                
                CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                
                 
               If LightOn <> 254 Then  ' bin5 : light on fail
                   rv1 = 1
                   rv2 = 1
                   rv3 = 1
                   rv4 = 2
                   AU6254TestMsg = "GPO R/W fail"
               End If
                Exit Sub
            End If
            
                


 End Sub
 
 Sub AU6254TestAlcorV2(rv0 As Byte, rv1 As Byte, rv2 As Byte, rv3 As Byte, rv4 As Byte)
 On Error Resume Next
 Dim TestCounter As Integer
 Dim TimeInterval
 Dim HubEmuCounter As Integer
 Dim ContinueRun As Integer
 Dim PwrTime As Single
 Dim Au6254Speed As Integer
 Dim ResetBit As Integer
 Dim FailPosition As Integer
 
 ContinueRun = 1
 PwrTime = 0.5
 ResetBit = 128
 
 txtmsg.Text = ""
              
                If PCI7248InitFinish = 0 Then
                  PCI7248ExistAU6254
                End If
                          
               '====================================
               '            Hub exist test3
               ' ====================================
               '****************************
                'For HubEmuCounter = 1 To 5
               '****************************
                    WinExec "off.exe", 0
                    CardResult = DO_WritePort(card, Channel_P1CL, &H4)
                '  CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                 ' Print "1"
                  Call MsecDelay(0.3) ' system let hub driver unload time
               
                  '============= dETECT FOR Disconnect fail
       
                  
                  Cls
                  Print "2"
                    WinExec "on.exe", 0
                  CardResult = DO_WritePort(card, Channel_P1A, &H1E + ResetBit)
                  Call MsecDelay(CSng(Text25.Text)) ' Hub Power on
                  If CardResult <> 0 Then
                      MsgBox "Power on fail"
                      End
                  End If
                  ReaderExist = 0
                  rv0 = 0
                  OldTimer = Timer
                  Cls
                  Print "3"
                  Do
                      Call MsecDelay(0.2)
                      DoEvents
                       rv0 = AU6254_GetDevice(0, 1, "6254")
                       
                      TimeInterval = Timer - OldTimer
                  Loop While rv0 = 0 And TimeInterval < 5
                 
                  Print "rv0 ="; rv0; "TimeInterval="; TimeInterval
                  
                 '****************************
              '  Next HubEmuCounter
              '  MsgBox "Ok"
              '  Exit Sub
                '****************************
                
                If rv0 <> 1 Then 'Hub unknow
                     rv1 = 4
                     rv2 = 4
                     rv3 = 4
                     rv4 = 4
                     If ContinueRun = 0 Then
                        Exit Sub
                     End If
                End If
        
                Print "rv0="; rv0; "--- Hub Exist"
    
                '==========================  Hub Detect
                '1. usb 2.0 reader
                '2. usb 1.0 flash
                '3, usb 6610
                '4. usb keyboard
        
               '======== PWR ctrl
               If ChipName <> "AU6254XLT20" Then
                CardResult = DO_WritePort(card, Channel_P1A, &H1C + ResetBit)
               End If
                Call MsecDelay(PwrTime)
                
                '===============   must at here to test speed , otherwise driver will overlap the test result
                
                HubPort = 0
                ReaderExist = 0
                rv1 = 0
                ClosePipe
                'rv1 = AU6610Test
                rv1 = CBWTest_New_AU9254(0, 1, "6362")
                ClosePipe
                Au6254Speed = UsbSpeedTestResult
                 
                Print "rv1="; rv1; "--- 2.0 speed test"
                
                
                   
               If rv1 <> 1 Then
                   
                    rv2 = 4
                    rv3 = 4
                    rv4 = 4
                    
                    
                        If rv1 >= 2 Then  ' speed error
                             Exit Sub
                        End If
                    
                    If ContinueRun = 0 Then
                        Exit Sub
                    End If
                End If
                
                 Print "rv1="; rv1; "--- 2.0 speed and isochrous pipe test"
               
                If ChipName <> "AU6254XLT20" Then
                CardResult = DO_WritePort(card, Channel_P1A, &H18 + ResetBit)
               
                Call MsecDelay(PwrTime)
                CardResult = DO_WritePort(card, Channel_P1A, &H10 + ResetBit)
                Call MsecDelay(PwrTime)
                CardResult = DO_WritePort(card, Channel_P1A, &H0 + ResetBit)
                Call MsecDelay(PwrTime)
                 
                Else
                
                CardResult = DO_WritePort(card, Channel_P1A, &H16 + ResetBit)
                Call MsecDelay(PwrTime)
                
                End If
                
                
                
                 
                '===== 6335 test usb 2.0 speed
              
                   
                 
                
                '****************************
              '  Next HubEmuCounter
              '  MsgBox "Ok"
              '  Exit Sub
                '****************************
             
                 
                 
               '=================================
               '   test usb 1.1 flash
               '================================
                HubPort = 1
                LBA = LBA + 1
                ReaderExist = 0
                ClosePipe
                  rv2 = CBWTest_New_AU9254(0, 1, "6362")
                ClosePipe
           
                Print "rv2="; rv2; "--- 2.0 speed and Bulk r/w"
                Print " UsbSpeedTestResult="; UsbSpeedTestResult; "---usb speed error r/w"
                
                
                If rv2 <> 1 Then
    
                   rv3 = 4
                   rv4 = 4
                   If ContinueRun = 0 Then
                        Exit Sub
                    End If
                End If
                 
                
               '=================================
               '   test 6610 for isochrous pipe
               '================================
               
               If ChipName = "AU6254XLT20" Then
               HubPort = 0
               Else
                HubPort = 2
               End If
               
                ReaderExist = 0
                ClosePipe
                rv3 = CBWTest_New_AU9254(0, 1, "9360")
                rv3 = CBWTest_New_8_Sector(0, 1, FailPosition, 20)
                ClosePipe
         
                Print "rv3="; rv3; "--- 1.1 speed and isochrorous r/w"
               
               
                 CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                 If rv3 <> 1 Then
                  rv4 = 4
                    If ContinueRun = 0 Then
                        Exit Sub
                    End If
                     
                  Else
                  
                     ' If LightON <> 243 Then
                     '     rv3 = 2
                      '   End If
                  
                     
                  End If
                
                
                 
                
                
                
               '=================================
               '   test usb 1.1 key board
               '================================
                  
                 txtmsg.Text = ""
                 txtmsg.SetFocus
                 
                 HubPort = 3
                 ReaderExist = 0
                 ClosePipe
                   rv4 = CBWTest_New_AU9254(0, 1, "413c")
                 ClosePipe
                
                
                   Print "rv4="; rv4; "--- 12 MHZ speed and interrupt  r/w"
                
                
                    
                 If rv0 = 1 Or rv1 = 1 Or rv2 = 1 Or rv3 = 1 Or rv4 = 1 Then
                 
                 
                    If rv0 * rv1 * rv2 * rv3 * rv4 = 1 Then
                      Rv1ContinueFail = 0
                    End If
                    
                  
                    If rv0 <> 1 Then
                    'Rv1ContinueFail = 0
                     ReaderExist = 0
                     rv0 = AU6254_GetDevice(0, 1, "6254")
                    End If
                    
                  If rv1 <> 1 Then
    
                    HubPort = 0
                    ReaderExist = 0
                    ClosePipe
                    'rv1 = AU6610Test
                    rv1 = CBWTest_New_AU9254(0, 1, "6362")
                    ClosePipe
                     Au6254Speed = UsbSpeedTestResult
     
                     
                     
                  End If
                  
                  
                  If rv2 <> 1 Then
                    ' Rv1ContinueFail = 0
                    
                     HubPort = 1
                     ReaderExist = 0
                    ClosePipe
                    rv2 = CBWTest_New_AU9254(0, 1, "6362")
                    ClosePipe
                  End If
                  
                  If rv3 <> 1 Then
                    'Rv1ContinueFail = 0
                    HubPort = 2
                    ReaderExist = 0
                      ClosePipe
                     rv3 = CBWTest_New_AU9254(0, 1, "9360")
                     rv3 = CBWTest_New_8_Sector(0, 1, FailPosition, 20)
                    ClosePipe
              
                  End If
                  
                  If rv4 <> 1 Then
                   '  Rv1ContinueFail = 0
                     HubPort = 3
                     ReaderExist = 0
                     ClosePipe
                     rv4 = CBWTest_New_AU9254(0, 1, "413c")
                     ClosePipe
        
                  End If
                  
                  
                    If rv1 + rv2 + rv3 + rv4 = 0 Then
                     
                      Rv1ContinueFail = Rv1ContinueFail + 1
                     End If
                      
                     
                     
                    If Rv1ContinueFail > 1 Then  ' use another AU6335 plug into to solve the hub hang probelm
                    CardResult = DO_WritePort(card, Channel_P1CL, &H0)
                    Call MsecDelay(2.5)
                    CardResult = DO_WritePort(card, Channel_P1CL, &H4)
                    Call MsecDelay(2)
                   End If
                 
                Print "RTrv0="; rv0
                Print "RTrv1="; rv1
                Print "RTrv2="; rv2
                Print "RTrv3="; rv3
                Print "RTrv4="; rv4

                
             End If
             
             
            '========== binning
             If rv0 = 0 Then   ' hub unknow  ,bin2
               rv1 = 4
               rv2 = 4
               rv3 = 4
               rv4 = 4
               AU6254TestMsg = "Hub Unknow"
               Exit Sub
          End If
          
        
               If rv0 = 1 Then   ' hub unknow  , bin3
                 If Au6254Speed = 2 Then     ' hub speed
                   rv1 = 2
                   rv2 = 4
                   rv3 = 4
                   rv4 = 4
                    AU6254TestMsg = "USB 2.0 speed error"
                   Exit Sub
                End If
                
                If rv1 * rv2 * rv3 * rv4 = 0 Then  ' bin4 down stream port unknow device
                     PortFail = ""
                     If rv1 = 0 Then
                       PortFail = PortFail & "port1 unknow,"
                    End If
                    
                       If rv2 = 0 Then
                       PortFail = PortFail & "port2 unknow,"
                    End If
                    
                       If rv3 = 0 Then
                       PortFail = PortFail & "port3 unknow,"
                    End If
                    
                       If rv4 = 0 Then
                       PortFail = PortFail & "port4 unknow,"
                    End If
                      
                      
                    
                    rv1 = 1
                    rv2 = 1
                    rv3 = 2
                    rv4 = 4
                    
                    AU6254TestMsg = PortFail
                    
                    Exit Sub
                End If
                
                
                
                
                If rv2 >= 2 Then   ' bin5 down stream port unknow device
                    rv1 = 1
                    rv2 = 1
                    rv3 = 1
                    rv4 = 2
                    AU6254TestMsg = "2.0 Reader SD R/W fail"
                    Exit Sub
                End If
                
                  
                If rv3 >= 2 Then   ' bin5 down stream port unknow device
                    rv1 = 1
                    rv2 = 1
                    rv3 = 1
                    rv4 = 2
                    AU6254TestMsg = "1.0 Reader SD R/W fail"
                    Exit Sub
                End If
                
                
                CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                
                 
               If LightOn <> 225 And LightOn <> 224 Then ' bin5 : light on fail
                   rv1 = 1
                   rv2 = 1
                   rv3 = 1
                   rv4 = 2
                   AU6254TestMsg = "GPO R/W fail"
               End If
                Exit Sub
            End If
            
                


 End Sub
Function OpenPipe6331() As Byte
Dim WritePathName As String
Dim ReadPathName As String

OpenPipe6331 = 0
'WritePathName = Left(DevicePathName, Len(DevicePathName) - 2) & "\PIPE0"   '
WritePathName = DevicePathName
'Debug.Print Lba
'Debug.Print "WritePathName="; WritePathName
WriteHandle6331 = CreateFile _
             (WritePathName, _
            GENERIC_READ Or GENERIC_WRITE, _
            (FILE_SHARE_READ Or FILE_SHARE_WRITE), _
             Security, _
             OPEN_EXISTING, _
             0&, _
            0)
'Debug.Print "write handle"; WriteHandle
If WriteHandle6331 = 0 Then
  OpenPipe6331 = 0
  Exit Function
End If
OpenPipe6331 = 1
End Function

Function CBWTest_New6331(Lun As Byte, PreSlotStatus As Byte, Vid_PID As String) As Byte
Dim i As Integer
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long

   CBWDataTransferLength = 1024
 
'   For i = 0 To CBWDataTransferLength - 1
    
'         ReadData(i) = 0

'   Next

    If PreSlotStatus <> 1 Then
        CBWTest_New6331 = 4
        Exit Function
    End If
    '========================================
   
    CBWTest_New6331 = 0
    If LBA > 25 * 1024 Then
        LBA = 0
    End If
    '========================================
     TmpString = ""
    If ReaderExist = 0 Then
        Do
            DoEvents
            Call MsecDelay(0.1)
            TimerCounter = TimerCounter + 1
            TmpString = GetDeviceNameMulti6331(Vid_PID)
        Loop While TmpString = "" And TimerCounter < 10
    End If
    '=======================================
    If ReaderExist = 0 And TmpString <> "" Then
      ReaderExist = 1
    End If
    '=======================================
    If ReaderExist = 0 And TmpString = "" Then
      CBWTest_New6331 = 0   ' no readerExist
      ReaderExist = 0
      Exit Function
    End If
    '=======================================
    If OpenPipe6331 = 0 Then
      CBWTest_New6331 = 2   ' Write fail
      Exit Function
    End If
    Print DevicePathName
    '======================================
    
     ' for unitSpeed
    Dim ret As Integer
    
    ret = fnScsi2usb_TestUnitReady(WriteHandle6331)
    ret = fnScsi2usb_TestUnitReady(WriteHandle6331)
    ret = fnScsi2usb_Write(WriteHandle6331, 1, LBA, Pattern(0))
    ret = fnScsi2usb_Read(WriteHandle6331, 1, LBA)
    
    Receive_DataBuffer
    
    Dim i2 As Integer
    
    For i2 = 1 To 500
        If Pattern(i2) <> ReceiveDataBuffer(i2 + 1) Then
        CBWTest_New6331 = 2
        Exit Function
        End If
    Next i2
    
     CBWTest_New6331 = 1
   
  End Function

 
 
Function CBWTest_New_AU6371(Lun As Byte, PreSlotStatus As Byte, Vid_PID As String) As Byte
Dim i As Integer
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long

   CBWDataTransferLength = 2048
 
'   For i = 0 To CBWDataTransferLength - 1
    
'         ReadData(i) = 0

'   Next

    If PreSlotStatus <> 1 Then
        CBWTest_New_AU6371 = 4
        Exit Function
    End If
    '========================================
   
    CBWTest_New_AU6371 = 0
    If LBA > 25 * 1024 Then
        LBA = 0
    End If
    '========================================
     TmpString = ""
    If ReaderExist = 0 Then
        Do
            DoEvents
            Call MsecDelay(0.1)
            TimerCounter = TimerCounter + 1
            TmpString = GetDeviceName(Vid_PID)
        Loop While TmpString = "" And TimerCounter < 10
    End If
    '=======================================
    If ReaderExist = 0 And TmpString <> "" Then
      ReaderExist = 1
    End If
    '=======================================
    If ReaderExist = 0 And TmpString = "" Then
      CBWTest_New_AU6371 = 0   ' no readerExist
      ReaderExist = 0
      Exit Function
    End If
    '=======================================
    If OpenPipe = 0 Then
      CBWTest_New_AU6371 = 2   ' Write fail
      Exit Function
    End If
 
    '======================================
    
     ' for unitSpeed
    
    TmpInteger = TestUnitSpeed(Lun)
    
    If TmpInteger = 0 Then
        
       CBWTest_New_AU6371 = 2   ' usb 2.0 high speed fail
       UsbSpeedTestResult = 2
       Exit Function
    End If
    
    
    
    TmpInteger = TestUnitReady(Lun)
    If TmpInteger = 0 Then
        TmpInteger = RequestSense(Lun)
        
        If TmpInteger = 0 Then
        
           CBWTest_New_AU6371 = 2  'Write fail
           Exit Function
        End If
        
    End If
    '======================================
    TmpInteger = Read_Data1(LBA, Lun, CBWDataTransferLength)
    
   ' TmpInteger = Read_Data1(Lba, Lun, CBWDataTransferLength)
    TmpInteger = Read_Data(LBA, Lun, CBWDataTransferLength)
      
    If TmpInteger = 0 Then
         CBWTest_New_AU6371 = 2  'write fail
          Exit Function
     End If
    
      
    TmpInteger = Write_Data(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
        CBWTest_New_AU6371 = 2  'write fail
        Exit Function
    End If
    
    TmpInteger = Read_Data(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
        CBWTest_New_AU6371 = 3    'Read fail
        Exit Function
    End If
     
    For i = 0 To CBWDataTransferLength - 1
    
        If ReadData(i) <> Pattern(i) Then
          CBWTest_New_AU6371 = 3    'Read fail
          Exit Function
        End If
    
    Next
    
    CBWTest_New_AU6371 = 1
        
    
    End Function
 
 Sub AU6254TestOld(rv0 As Byte, rv1 As Byte, rv2 As Byte, rv3 As Byte, rv4 As Byte)
 

   
              
                If PCI7248InitFinish = 0 Then
                  PCI7248ExistAU6254
                End If
                          
               '====================================
               '            Hub exist test
               ' ====================================
                CardResult = DO_WritePort(card, Channel_P1A, &H1E)   ' Hub Power on
                If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                End If
                ReaderExist = 0
                HubPort = 0
                Call MsecDelay(2.8)    'power on time
                rv0 = AU6254_GetDevice(0, 1, "6254")
                 
                  
                 If rv0 <> 1 Then
                    rv1 = 4
                    rv2 = 4
                    rv3 = 4
                    rv4 = 4
                    Exit Sub
                 End If
        
               
                '==========================  Hub Detect
                '1. usb 2.0 reader
                '2. usb 1.0 flash
                '3, usb 6610
                '4. usb keyboard
                 
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H1C)   ' usb 2.0 falsh
                
                Call MsecDelay(0.8)    'power on time
                 
                If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                End If
                
                
                ' test usb 2.0 speed
                HubPort = 0
                ReaderExist = 0
                ClosePipe
                rv1 = CBWTest_New(0, 1, "9380")
                ClosePipe
                
                
                If rv1 <> 1 Then
                    
                    rv2 = 4
                    rv3 = 4
                    rv4 = 4
                     Exit Sub
                End If
                
                
                
               '=================================
               '   test usb 1.1 flash
               '================================
                
        
                 HubPort = 1
                 ReaderExist = 0
                ClosePipe
               '    rv2 = CBWTest_New(0, 1, "55AA")
                ClosePipe
                
                 rv2 = 1
                If rv2 <> 1 Then
             
                    rv3 = 4
                    rv4 = 4
                    'Exit Sub
                End If
                
                  Print "rv2="; rv2
               '=================================
               '   test 6610 for isochrous pipe
               '================================
                 
                
                CardResult = DO_WritePort(card, Channel_P1A, &H10)   ' usb 6610
                
                Call MsecDelay(0.8)    'power on time
                 
                If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                End If
                
                rv3 = AU6610Test
                Print "rv3="; rv3
                    
               '=================================
               '   test usb 1.1 key board
               '================================
                
                 txtmsg.Text = ""
                 txtmsg.SetFocus
        
                 CardResult = DO_WritePort(card, Channel_P1A, &H0)    ' usb 1.0 keyboard
                
                Call MsecDelay(4)    'power on time
                 
                If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                End If
                
                ' keybaord control
                
                 
                 
                  CardResult = DO_WritePort(card, Channel_P1CH, &H1)
                    Call MsecDelay(0.5)
                
                  CardResult = DO_WritePort(card, Channel_P1CH, &H2)
                    Call MsecDelay(0.5)
                    
                  CardResult = DO_WritePort(card, Channel_P1CH, &H4)
                    Call MsecDelay(0.5)
                    
                  CardResult = DO_WritePort(card, Channel_P1CH, &H8)
                    Call MsecDelay(0.5)
                    
                    
                  CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' close power
                   
                  Print txtmsg.Text
                  
                  If InStr(txtmsg.Text, "g=") <> "" Then
                    rv4 = 1
                  End If
                  

 End Sub
 Sub AU6254TestLoop(rv0 As Byte, rv1 As Byte, rv2 As Byte, rv3 As Byte, rv4 As Byte)
 On Error Resume Next
 Dim TestCounter As Integer
 Dim TimeInterval
 Dim HubEmuCounter As Integer
 txtmsg.Text = ""
              
                If PCI7248InitFinish = 0 Then
                  PCI7248ExistAU6254
                End If
                          
               '====================================
               '            Hub exist test3
               ' ====================================
               For HubEmuCounter = 1 To 10
                 
                  CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                  Call MsecDelay(4)
                  CardResult = DO_WritePort(card, Channel_P1A, &H1E)
                  Call MsecDelay(0.5) ' Hub Power on
                  If CardResult <> 0 Then
                      MsgBox "Power on fail"
                      End
                  End If
                  ReaderExist = 0
                  OldTimer = Timer
                   rv0 = 0
                  Do
                      Call MsecDelay(0.2)
                      DoEvents
                      rv0 = AU6254_GetDevice(0, 1, "6254")
                      TimeInterval = Timer - OldTimer
                  Loop While rv0 = 0 And TimeInterval < 15
                 
                  Print "rv0 ="; rv0; "TimeInterval="; TimeInterval
                Next HubEmuCounter
                
                MsgBox "Ok"
                Exit Sub
                  If rv0 <> 1 Then
                     rv1 = 4
                     rv2 = 4
                     rv3 = 4
                     rv4 = 4
                    ' Exit Sub
                  End If
        
                 Print "rv0="; rv0; "--- Hub Exist"
             
                 
                
                '==========================  Hub Detect
                '1. usb 2.0 reader
                '2. usb 1.0 flash
                '3, usb 6610
                '4. usb keyboard
        
                
                ' test usb 2.0 speed
                 HubPort = 0
                 
                ReaderExist = 0
                
                
                
                ClosePipe
                    'rv1 = AU6610Test
                   rv1 = CBWTest_New_AU9254(0, 1, "6335")
                  
                ClosePipe
                 
                
                
                 
                 Print "rv1="; rv1; "--- 2.0 speed and isochrous pipe test"
                If rv1 <> 1 Then
                    If rv1 = 0 Then
                    rv0 = 0
                    rv1 = 4
                    End If
                     
                    rv2 = 4
                    rv3 = 4
                    rv4 = 4
                   '   Exit Sub
                End If
                
                 Print "rv1="; rv1; "--- 2.0 speed and isochrous pipe test"
                 
                 
               '=================================
               '   test usb 1.1 flash
               '================================
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H18)     ' usb 2.0 falsh
               
                Call MsecDelay(2.8)
                
                 LBA = LBA + 1
                ReaderExist = 0
                ClosePipe6331
                 ' rv2 = CBWTest_New_AU9254(0, rv1, "6331")
                 rv2 = CBWTest_New6331(0, 1, "6331")
                ClosePipe6331
                  
                    Dim tp As String
            '   tp = Dir("F:\*.*")
               
            '   If InStr(tp, "6254") <> 0 Then  ' bulk transfer
            '      rv2 = 1
            '   End If
                  
               ' rv2 = 1
                
                Print "rv2="; rv2; "--- 2.0 speed and Bulk r/w"
                Print " UsbSpeedTestResult="; UsbSpeedTestResult; "---usb speed error r/w"
                 Print " file name 0f 6331="; tp; "---usb speed error r/w"
                If rv2 <> 1 Then
                
                   If rv2 = 0 Then
                   rv2 = 2
                   End If
             
                    rv3 = 4
                    rv4 = 4
                    ' Exit Sub
                End If
                 
                
               '=================================
               '   test 6610 for isochrous pipe
               '================================
               
               
                CardResult = DO_WritePort(card, Channel_P1A, &H0)     ' usb 2.0 falsh
               
                Call MsecDelay(1)
               
                 HubPort = 1
                ReaderExist = 0
                ClosePipe
                  rv3 = CBWTest_New_AU9254(0, 1, "9369")
                ClosePipe
             
             
                 Print "rv3="; rv3; "--- 1.1 speed and isochrorous r/w"
               
               
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                 If rv3 <> 1 Then
                        If rv3 = 0 Then
                          rv3 = 2
                        End If
                    
                       
                      rv4 = 4
                   '  Exit Sub
                     
                  Else
                  
                     ' If LightON <> 243 Then
                     '     rv3 = 2
                      '   End If
                  
                     
                End If
                
                
                 
                
                
                
               '=================================
               '   test usb 1.1 key board
               '================================
               ' CardResult = DO_WritePort(card, Channel_P1A, &HE)
               ' Call MsecDelay(3)
               '  rv3 = 1
                 txtmsg.Text = ""
                 txtmsg.SetFocus
                 
                 HubPort = 2
                 ReaderExist = 0
                 ClosePipe
                   rv4 = CBWTest_New_AU9254(0, 1, "9462")
                 ClosePipe
                
                ' ReaderExist = 0
                ' ClosePipe
                '   rv4 = CBWTest_New_AU9254(0, 1, "9462")
                ' ClosePipe
        
                  '  CardResult = DO_WritePort(card, Channel_P1A, &HE)    ' usb 2.0 falsh
                
              '  Call MsecDelay(4)
                
                ' keybaord control
                 '  CardResult = DO_ReadPort(card, Channel_P1B, LightON)
                 '  Call MsecDelay(0.1)
                '   CardResult = DO_WritePort(card, Channel_P1CH, &H0)
                 '   Call MsecDelay(0.3)
                 
                 ' CardResult = DO_WritePort(card, Channel_P1CH, &H1)
                  '    Call MsecDelay(1.2)
                
                 
                    
                 ' CardResult = DO_WritePort(card, Channel_P1CH, &H0)
                 '   Call MsecDelay(0.2)
                  'CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' close power
                   
                 ' Print " LightON="; LightON
                  
                  
                   
                   
                   
               '    If InStr(txtmsg.Text, "..") <> 0 Then
                  
                '      rv4 = 1
                    
                '      If LightON = 238 Then
                '      rv4 = 1
                     
                 '    Else
                 '     Print "GPO Fail"
                     
                  '   End If
                     
               '  Else
               '    Print "Keyboard fail"
                    
                 
               '  End If
                  
                   
                   
                  
                   
                  
                  If rv4 <> 1 Then
                      rv4 = 2
                  End If
                Print "rv4="; rv4; "--- 12 MHZ speed and interrupt  r/w"
                
                
                    If rv0 = 1 Or rv1 = 1 Or rv2 = 1 Or rv3 = 1 Or rv4 = 1 Then
                  
                  If rv0 <> 1 Then
                   rv0 = AU6254_GetDevice(0, 1, "6254")
                  End If
                  
                  If rv1 <> 1 Then
                  
                    HubPort = 0
                    ReaderExist = 0
                    ClosePipe
                    'rv1 = AU6610Test
                    rv1 = CBWTest_New_AU9254(0, 1, "6335")
                    ClosePipe
                    
                  End If
                  
                  
                  If rv2 <> 1 Then
                    ClosePipe6331
                    rv2 = CBWTest_New6331(0, 1, "6331")
                    ClosePipe6331
                  End If
                  
                  If rv3 <> 1 Then
                    HubPort = 1
                    ReaderExist = 0
                    ClosePipe
                    rv3 = CBWTest_New_AU9254(0, 1, "9369")
                    ClosePipe
                  End If
                  
                  If rv4 <> 1 Then
                  
                     HubPort = 2
                     ReaderExist = 0
                    ClosePipe
                    rv4 = CBWTest_New_AU9254(0, 1, "9462")
                    ClosePipe
        
                  End If
              
             Print "rv0="; rv0
             Print "rv1="; rv1
             Print "rv2="; rv2
             Print "rv3="; rv3
             Print "rv4="; rv4

                
             End If
                


 End Sub
  
 Sub AU6375ASTest(rv0 As Byte, rv1 As Byte, rv2 As Byte)

 
 
                ClosePipe
                If ChipName = "AU6375AS" Then

                rv0 = AU6375_GetDevice(0, 1, "6362")
                End If
                
                
                If ChipName = "AU6371AS" Then

                rv0 = AU6375_GetDevice(0, 1, "6387")
                End If
                
                  If ChipName = "AUHang" Then

                rv0 = AU6375_GetDevice(0, 1, "vid")
                End If
                
                
                ClosePipe
                If rv0 = 1 Then
                    OpenPipe
                    Print "testmode"
                     Print "Old Hang Time="; AU6375ASHangTime; " s"
                     If ChipName <> "AUHang" Then
                      rv1 = AU6375TestMode(0)
                     Else
                      rv1 = 1
                     End If
                     
                    ClosePipe
                    If rv1 = 1 Then
                        OpenPipe
                        Print "Endless"
                        
                        rv2 = AU6375EndLess(0)
                        
                        AU6375ASHangTime = Timer - OldTimer
                        Print "Hang Time="; AU6375ASHangTime; " s"
                        If AU6375ASHangTime > 3 Then
                            rv2 = 1   '============ hang success
                        Else
                            rv2 = 2
                        End If
                        
                        ' for system hang,
                        ' if fail chip, system no hang, host get rv2=2
                        ' if pass chip, system hang, host get nothing
                        
                        ClosePipe
                        Print "Endless over"
                    Else
                        rv1 = 2 ' rv1 fail
                        rv2 = 4

                    End If
          
                     
                Else
                    rv1 = 4  ' rv0 fail
                    rv2 = 4
                     
               End If
               
               
               
               
 End Sub
 
 Sub AU6375ASTestOld(rv0 As Byte, rv1 As Byte, rv2 As Byte)

 
 
                ClosePipe
                If ChipName = "AU6375AS" Then

                rv0 = AU6375_GetDevice(0, 1, "6362")
                End If
                
                
                If ChipName = "AU6371AS" Then

                rv0 = AU6375_GetDevice(0, 1, "6366")
                End If
                
                ClosePipe
                If rv0 = 1 Then
                    OpenPipe
                    Print "testmode"
                     Print "Old Hang Time="; AU6375ASHangTime; " s"
                    rv1 = AU6375TestMode(0)
                    ClosePipe
                    If rv1 = 1 Then
                        OpenPipe
                        Print "Endless"
                        
                        rv2 = AU6375EndLess(0)
                        
                        AU6375ASHangTime = Timer - OldTimer
                        Print "Hang Time="; AU6375ASHangTime; " s"
                        If AU6375ASHangTime > 3 Then
                            rv2 = 1   '============ hang success
                        Else
                            rv2 = 2
                        End If
                        
                        ' for system hang,
                        ' if fail chip, system no hang, host get rv2=2
                        ' if pass chip, system hang, host get nothing
                        
                        ClosePipe
                        Print "Endless over"
                    Else
                        rv1 = 2 ' rv1 fail
                        rv2 = 4

                    End If
          
                     
                Else
                    rv1 = 4  ' rv0 fail
                    rv2 = 4
                     
               End If
               
               
               
               
 End Sub
 
 
Public Function AU6610TestRC() As Integer
On Error Resume Next
AU6610TestRC = Test_RC

End Function
Function AU9520_GetDeviceName(Lun As Byte, PreSlotStatus As Byte, Vid_PID As String) As Byte
Dim i As Integer
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long

 

TmpString = ""
If ReaderExist = 0 Then
Do
DoEvents
 
               Call MsecDelay(0.1)
               TimerCounter = TimerCounter + 1
              TmpString = GetDeviceName(Vid_PID)
  
Loop While TmpString = "" And TimerCounter < 10
End If
If ReaderExist = 0 And TmpString <> "" Then
  ReaderExist = 1
End If

If ReaderExist = 0 And TmpString = "" Then
  AU9520_GetDeviceName = 0   ' no readerExist
  ReaderExist = 0
  Exit Function
End If
AU9520_GetDeviceName = 1
End Function

 

Function STDHex(VarReadData As Byte) As String
 
STDHex = Trim(CStr(Hex(VarReadData)))
If Len(STDHex) < 2 Then
STDHex = "0" & STDHex
End If




End Function


Function DumpTable_AU6375A31(SectorNo As Byte, CBWDataTransferLength As Long) As Byte
Dim CBW(0 To 30) As Byte
Dim NumberOfBytesWritten As Long
Dim CBWDataTransferLen(0 To 3) As Byte
  
Dim TransferLen As Long
Dim TransferLenLSB As Byte
Dim TransferLenMSB As Byte
Dim i As Integer
Dim tmpV(0 To 2) As Long
Dim opcode As Byte

Dim CSW(0 To 12) As Byte

Dim NumberOfBytesRead As Long

For i = 0 To 30
   
        CBW(i) = 0
    
Next i

For i = 0 To CBWDataTransferLength
ReadData(i) = 0
Next

Const CBWSignature_0 = &H55
Const CBWSignature_1 = &H53
Const CBWSignature_2 = &H42
Const CBWSignature_3 = &H43


Const CBWTag_0 = &H1
Const CBWTag_1 = &H2
Const CBWTag_2 = &H3
Const CBWTag_3 = &H4


'/////////////////// CBW signature

CBW(0) = CBWSignature_0
CBW(1) = CBWSignature_1
CBW(2) = CBWSignature_2
CBW(3) = CBWSignature_3

'/////////////////  CBW Tag

CBW(4) = CBWTag_0
CBW(5) = CBWTag_1
CBW(6) = CBWTag_2
CBW(7) = CBWTag_3

CBWDataTransferLen(0) = (CBWDataTransferLength Mod 256)
tmpV(0) = Int(CBWDataTransferLength / 256)
CBWDataTransferLen(1) = (tmpV(0) Mod 256)
tmpV(1) = Int(tmpV(0) / 256)
CBWDataTransferLen(2) = (tmpV(1) Mod 256)
tmpV(2) = Int((tmpV(1) / 256))
CBWDataTransferLen(3) = (tmpV(2) Mod 256)

CBW(8) = CBWDataTransferLen(0)  '00
CBW(9) = CBWDataTransferLen(1)  '08
CBW(10) = CBWDataTransferLen(2) '00
CBW(11) = CBWDataTransferLen(3) '00

'///////////////  CBW Flag
CBW(12) = &H80                 '80

'////////////// LUN
CBW(13) = 0                   '00

'///////////// CBD Len
CBW(14) = &H8               '0a

'////////////  UFI command

CBW(15) = &HC7
CBW(16) = &H12
 

CBW(17) = 0                 '00
CBW(18) = SectorNo         '00
CBW(19) = 0        '00
CBW(20) = 0         '40

'/////////////  Reverve
CBW(21) = 0

'//////////// Transfer Len

 

CBW(22) = 0    '00
CBW(23) = 0    '04

For i = 24 To 30
    CBW(i) = 0
Next

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
 
Dim result As Long

'1. CBW command

 
result = WriteFile _
       (WriteHandle, _
       CBW(0), _
       31, _
       NumberOfBytesWritten, _
       0)    'out

If result = 0 Then
 DumpTable_AU6375A31 = 0
 Exit Function
End If

'2. Readdata stage
 
result = ReadFile _
         (ReadHandle, _
          ReadData(0), _
          CBWDataTransferLength, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in

 
If result = 0 Then
 DumpTable_AU6375A31 = 0
 Exit Function
End If

'3. CSW data
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 
If result = 0 Then
 DumpTable_AU6375A31 = 0
 Exit Function
End If
 
'4. CSW status


If CSW(12) = 1 Then
    DumpTable_AU6375A31 = 0
Else
    DumpTable_AU6375A31 = 1
    
    
   
End If

End Function

Function DumpTableTest2(Table As Byte) As Byte
Dim i As Integer
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long
Dim TmpValue As Byte
Dim TableStr As String
Dim tmpStr As String

   CBWDataTransferLength = 512
 
    For i = 0 To CBWDataTransferLength - 1
    
          ReadData(i) = 0

    Next

   
    '========================================
   
    DumpTableTest2 = 0
    
    '========================================
     TmpString = ""
    If ReaderExist = 0 Then
        Do
            DoEvents
            Call MsecDelay(0.1)
            TimerCounter = TimerCounter + 1
            TmpString = GetDeviceName("058f")
        Loop While TmpString = "" And TimerCounter < 10
    End If
    '=======================================
    If ReaderExist = 0 And TmpString <> "" Then
      ReaderExist = 1
    End If
    '=======================================
    If ReaderExist = 0 And TmpString = "" Then
      DumpTableTest2 = 0   ' no readerExist
      ReaderExist = 0
      Exit Function
    End If
    '=======================================
    If OpenPipe = 0 Then
      DumpTableTest2 = 2   ' Write fail
      Exit Function
    End If
  
     '======================================
    TmpInteger = TestUnitReady(2)   ' XD slot
    If TmpInteger = 0 Then
        TmpInteger = RequestSense(2)
        
        If TmpInteger = 0 Then
        
           DumpTableTest2 = 2  'Write fail
           Exit Function
        End If
        
    End If
    
    TmpInteger = Read_Data(2, 2, CBWDataTransferLength)
  '  TmpInteger = Read_Data(2, 2, CBWDataTransferLength) ' for card chanage error
    If TmpInteger = 0 Then
        DumpTableTest2 = 2  'write fail
        Exit Function
    End If
    
    
     '====================================== dump buffer 3
 '   TableStr = ""
    TmpValue = DumpTable_AU6375A31(4, CBWDataTransferLength)
    
    TableStr = TableStr & "BufferNo = 4" & vbCrLf
  
    For i = 0 To 511
        tmpStr = STDHex(ReadData(i))
        TableStr = TableStr & tmpStr & " "
    
    
       If Table = 0 Then
            If DumpTableNormal(i + 512 * 2) <> tmpStr Then
               DumpTableTest2 = 33
               Exit Function
             End If
       Else
             If DumpTableInverse(i + 512 * 2) <> tmpStr Then
               DumpTableTest2 = 33
               Exit Function
             End If
       End If
          
        
           If ((i + 1) Mod 16) = 0 Then
             TableStr = TableStr & vbCrLf
        
           End If
    
    Next i
  
      'TableShow.Text1.Text = TableStr
    
    '====================================== dump buffer 5
   
    TmpValue = DumpTable_AU6375A31(5, CBWDataTransferLength)
    
    TableStr = TableStr & "BufferNo = 5" & vbCrLf
  
    For i = 0 To 511
        tmpStr = STDHex(ReadData(i))
        TableStr = TableStr & tmpStr & " "
    
       If Table = 0 Then
            If DumpTableNormal(i + 512 * 3) <> tmpStr Then
              DumpTableTest2 = 34
              Exit Function
            End If
       Else
            If DumpTableInverse(i + 512 * 3) <> tmpStr Then
              DumpTableTest2 = 34
              Exit Function
            End If
       End If

        
        
           If ((i + 1) Mod 16) = 0 Then
             TableStr = TableStr & vbCrLf
        
           End If
    
    Next i
  
    '====================================== dump buffer 6
 
    TmpValue = DumpTable_AU6375A31(6, CBWDataTransferLength)
    
    TableStr = TableStr & "BufferNo = 6" & vbCrLf
  
    For i = 0 To 511
        tmpStr = STDHex(ReadData(i))
        TableStr = TableStr & tmpStr & " "
    
    
       If Table = 0 Then
            If DumpTableNormal(i + 512 * 4) <> tmpStr Then
              DumpTableTest2 = 35
              Exit Function
            End If
      Else
             If DumpTableInverse(i + 512 * 4) <> tmpStr Then
               DumpTableTest2 = 35
               Exit Function
             End If
      End If
        
           If ((i + 1) Mod 16) = 0 Then
             TableStr = TableStr & vbCrLf
        
           End If
    
    Next i
  
  
  ' TableShow.Text1.Text = TableStr
    
   DumpTableTest2 = 1

  End Function

Function DumpTableTest(Table As Byte) As Byte
Dim i As Integer
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long
Dim TmpValue As Byte
Dim TableStr As String
Dim tmpStr As String

   CBWDataTransferLength = 512
 
'   For i = 0 To CBWDataTransferLength - 1
    
'         ReadData(i) = 0

'   Next

   
    '========================================
   
    DumpTableTest = 0
    
    '========================================
     TmpString = ""
    If ReaderExist = 0 Then
        Do
            DoEvents
            Call MsecDelay(0.1)
            TimerCounter = TimerCounter + 1
            TmpString = GetDeviceName("058f")
        Loop While TmpString = "" And TimerCounter < 10
    End If
    '=======================================
    If ReaderExist = 0 And TmpString <> "" Then
      ReaderExist = 1
    End If
    '=======================================
    If ReaderExist = 0 And TmpString = "" Then
      DumpTableTest = 0   ' no readerExist
      ReaderExist = 0
      Exit Function
    End If
    '=======================================
    If OpenPipe = 0 Then
      DumpTableTest = 2   ' Write fail
      Exit Function
    End If
   
    
    
       TmpInteger = TestUnitReady(3)   ' MS slot
        If TmpInteger = 0 Then
        TmpInteger = RequestSense(3)
        
        If TmpInteger = 0 Then
        
           DumpTableTest = 2  'Write fail
           Exit Function
        End If
        
    End If
    
   
    '====================================== dump buffer 2
    TableStr = ""
    TmpValue = DumpTable_AU6375A31(2, CBWDataTransferLength)
    
    TableStr = "BufferNo = 2" & vbCrLf
  
    For i = 0 To 511
        tmpStr = STDHex(ReadData(i))
        TableStr = TableStr & tmpStr & " "
    
      If Table = 0 Then
         If DumpTableNormal(i) <> tmpStr Then
           DumpTableTest = 31                   ' read fail , table fail
          Exit Function
         End If
      Else
         
         If DumpTableInverse(i) <> tmpStr Then
           DumpTableTest = 31                   ' read fail , table fail
          Exit Function
         End If
      End If
        
           If ((i + 1) Mod 16) = 0 Then
             TableStr = TableStr & vbCrLf
        
           End If
    
    Next i
     
     
  
   '====================================== dump buffer 3
   
    TmpValue = DumpTable_AU6375A31(3, CBWDataTransferLength)
    
    TableStr = TableStr & "BufferNo = 3" & vbCrLf
  
    For i = 0 To 511
        tmpStr = STDHex(ReadData(i))
        TableStr = TableStr & tmpStr & " "
    
      If Table = 0 Then
        If DumpTableNormal(i + 512) <> tmpStr Then
          DumpTableTest = 32
          Exit Function
        End If
        
     Else
        If DumpTableInverse(i + 512) <> tmpStr Then
          DumpTableTest = 32
          Exit Function
        End If
        
    End If
        
           If ((i + 1) Mod 16) = 0 Then
             TableStr = TableStr & vbCrLf
        
           End If
    
    Next i
  
  
  
     '======================================
    TmpInteger = TestUnitReady(2)    ' XD slot
    If TmpInteger = 0 Then
        TmpInteger = RequestSense(2)
        
        If TmpInteger = 0 Then
        
           DumpTableTest = 2  'Write fail
           Exit Function
        End If
        
    End If
    
    TmpInteger = Read_Data(2, 2, CBWDataTransferLength)
      
    If TmpInteger = 0 Then
        DumpTableTest = 2  'write fail
        Exit Function
    End If
    
    
     '====================================== dump buffer 3
 '   TableStr = ""
    TmpValue = DumpTable_AU6375A31(4, CBWDataTransferLength)
    
    TableStr = TableStr & "BufferNo = 4" & vbCrLf
  
    For i = 0 To 511
        tmpStr = STDHex(ReadData(i))
        TableStr = TableStr & tmpStr & " "
    
    
       If Table = 0 Then
            If DumpTableNormal(i + 512 * 2) <> tmpStr Then
               DumpTableTest = 33
               Exit Function
             End If
       Else
             If DumpTableInverse(i + 512 * 2) <> tmpStr Then
               DumpTableTest = 33
               Exit Function
             End If
       End If
          
        
           If ((i + 1) Mod 16) = 0 Then
             TableStr = TableStr & vbCrLf
        
           End If
    
    Next i
  
      'TableShow.Text1.Text = TableStr
    
    '====================================== dump buffer 5
   
    TmpValue = DumpTable_AU6375A31(5, CBWDataTransferLength)
    
    TableStr = TableStr & "BufferNo = 5" & vbCrLf
  
    For i = 0 To 511
        tmpStr = STDHex(ReadData(i))
        TableStr = TableStr & tmpStr & " "
    
       If Table = 0 Then
            If DumpTableNormal(i + 512 * 3) <> tmpStr Then
              DumpTableTest = 34
              Exit Function
            End If
       Else
            If DumpTableInverse(i + 512 * 3) <> tmpStr Then
              DumpTableTest = 34
              Exit Function
            End If
       End If

        
        
           If ((i + 1) Mod 16) = 0 Then
             TableStr = TableStr & vbCrLf
        
           End If
    
    Next i
  
    '====================================== dump buffer 6
 
    TmpValue = DumpTable_AU6375A31(6, CBWDataTransferLength)
    
    TableStr = TableStr & "BufferNo = 6" & vbCrLf
  
    For i = 0 To 511
        tmpStr = STDHex(ReadData(i))
        TableStr = TableStr & tmpStr & " "
    
    
       If Table = 0 Then
            If DumpTableNormal(i + 512 * 4) <> tmpStr Then
              DumpTableTest = 35
              Exit Function
            End If
      Else
             If DumpTableInverse(i + 512 * 4) <> tmpStr Then
               DumpTableTest = 35
               Exit Function
             End If
      End If
        
           If ((i + 1) Mod 16) = 0 Then
             TableStr = TableStr & vbCrLf
        
           End If
    
    Next i
  
  
   'TableShow.Text1.Text = TableStr
    
   DumpTableTest = 1

  End Function
Sub ReadTable()
Dim i As Integer
Dim tmp As String
Dim Str As String
Dim j As Integer

' ===================== read normal table
i = 0
Open App.Path & "\DumpTableNormal.txt" For Input As #2

    Do While Not EOF(2)
    
        Input #2, tmp
        If Len(tmp) > 20 Then
           ' Str = ""
            For j = 1 To 16
            
            DumpTableNormal(i) = Mid(tmp, (j - 1) * 3 + 1, 2)
           ' Str = Str & DumpTableNormal(i)
            i = i + 1
            Next j
        End If
       ' Debug.Print Str
          
        
    Loop

Close #2
'MsgBox i
' ===================== read inverse table

i = 0
Open App.Path & "\DumpTableInverse.txt" For Input As #2

     Do While Not EOF(2)
    
        Input #2, tmp
        If Len(tmp) > 20 Then
            Str = ""
            For j = 1 To 16
            
            DumpTableInverse(i) = Mid(tmp, (j - 1) * 3 + 1, 2)
           Str = Str & DumpTableInverse(i)
            i = i + 1
            Next j
        End If
      ' Debug.Print Str
          
        
    Loop

Close #2


'MsgBox i
' ===================== read inverse table


End Sub



Function Write_Firmware_Data(LBA As Long, Lun As Byte, CBWDataTransferLength As Long) As Byte

Dim CBW(0 To 30) As Byte
Dim CSW(0 To 12) As Byte
Dim NumberOfBytesWritten As Long
Dim NumberOfBytesRead As Long
Dim CBWDataTransferLen(0 To 3) As Byte
Dim TransferLen As Long
Dim TransferLenLSB As Byte
Dim TransferLenMSB As Byte
Dim i As Integer
Dim tmpV(0 To 2) As Long
Dim opcode As Byte

opcode = &H2A
'Buffer(0) = &H33 'CByte(Text2.Text)
'Buffer(1) = &H44


    For i = 0 To 30
    
        CBW(i) = 0
    
    Next i
    
Const CBWSignature_0 = &H55
Const CBWSignature_1 = &H53
Const CBWSignature_2 = &H42
Const CBWSignature_3 = &H43


Const CBWTag_0 = &H1
Const CBWTag_1 = &H2
Const CBWTag_2 = &H3
Const CBWTag_3 = &H4


'/////////////////// CBW signature

CBW(0) = CBWSignature_0
CBW(1) = CBWSignature_1
CBW(2) = CBWSignature_2
CBW(3) = CBWSignature_3

'/////////////////  CBW Tag

CBW(4) = CBWTag_0
CBW(5) = CBWTag_1
CBW(6) = CBWTag_2
CBW(7) = CBWTag_3

CBWDataTransferLen(0) = (CBWDataTransferLength Mod 256)
tmpV(0) = Int(CBWDataTransferLength / 256)
CBWDataTransferLen(1) = (tmpV(0) Mod 256)
tmpV(1) = Int(tmpV(0) / 256)
CBWDataTransferLen(2) = (tmpV(1) Mod 256)
tmpV(2) = Int((tmpV(1) / 256))
CBWDataTransferLen(3) = (tmpV(2) Mod 256)

CBW(8) = CBWDataTransferLen(0)  '00
CBW(9) = CBWDataTransferLen(1)  '08
CBW(10) = CBWDataTransferLen(2) '00
CBW(11) = CBWDataTransferLen(3) '00

'///////////////  CBW Flag
CBW(12) = &H0                 '80

'////////////// LUN
CBW(13) = Lun                    '00

'///////////// CBD Len
CBW(14) = &HA                '0a

'////////////  UFI command

CBW(15) = &HF0
CBW(16) = &H80
CBW(17) = &HF1
CBW(18) = &H5
CBW(19) = &H0
CBW(20) = &H0
CBW(21) = &H0
CBW(22) = &H0
CBW(23) = &H0
CBW(24) = &HF0
CBW(25) = &H10
CBW(26) = &HF4
CBW(27) = &H1
CBW(28) = &HF0
CBW(29) = &H70
CBW(30) = &HF3
CBW(31) = &H1


'1. CBW output
 
result = WriteFile _
       (WriteHandle, _
       CBW(0), _
       31, _
       NumberOfBytesWritten, _
       0)    'out

If result = 0 Then
   ' Write_Data = 0
    Exit Function
End If
 
 
 
'2, Output data
result = WriteFile _
       (WriteHandle, _
       Pattern(0), _
       CBWDataTransferLength, _
       NumberOfBytesWritten, _
       0)    'out

 
If result = 0 Then
   ' Write_Data = 0
    Exit Function
End If

'3 . CSW
result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
        
If result = 0 Then
    'Write_Data = 0
    Exit Function
End If
 
 
 
If CSW(12) = 1 Then
'Write_Data = 0

Else
'Write_Data = 1
End If
End Function

Function InquiryScSiString(Lun As Byte, CBWDataTransferLength As Long, ScSiRecMode As Byte) As Byte
Dim CBW(0 To 30) As Byte
Dim NumberOfBytesWritten As Long
Dim CBWDataTransferLen(0 To 3) As Byte
  
Dim TransferLen As Long
Dim TransferLenLSB As Byte
Dim TransferLenMSB As Byte
Dim i As Integer
Dim tmpV(0 To 2) As Long
Dim opcode As Byte

Dim CSW(0 To 12) As Byte

Dim NumberOfBytesRead As Long

'Dim Capacity(0 To 7) As Byte

 
For i = 0 To 30
   
        CBW(i) = 0
    
Next i

For i = 0 To CBWDataTransferLength
ReadData(i) = 0
Next

Const CBWSignature_0 = &H55
Const CBWSignature_1 = &H53
Const CBWSignature_2 = &H42
Const CBWSignature_3 = &H43


Const CBWTag_0 = &H1
Const CBWTag_1 = &H2
Const CBWTag_2 = &H3
Const CBWTag_3 = &H4


'/////////////////// CBW signature

CBW(0) = CBWSignature_0
CBW(1) = CBWSignature_1
CBW(2) = CBWSignature_2
CBW(3) = CBWSignature_3

'/////////////////  CBW Tag

CBW(4) = CBWTag_0
CBW(5) = CBWTag_1
CBW(6) = CBWTag_2
CBW(7) = CBWTag_3

 
CBW(8) = &H24  '00
CBW(9) = &H0  '08
CBW(10) = &H0 '00
CBW(11) = &H0 '00

'///////////////  CBW Flag
CBW(12) = &H80                 '80

'////////////// LUN
CBW(13) = Lun                    '00

'///////////// CBD Len
CBW(14) = &H6               '0a

'////////////  UFI command

CBW(15) = &H12
CBW(16) = Lun * 32
 
CBW(17) = &H0         '00
CBW(18) = &H0        '00
CBW(19) = &H24       '00
CBW(20) = &H0         '40

'/////////////  Reverve
CBW(21) = 0

'//////////// Transfer Len

 
CBW(22) = &H0     '00
CBW(23) = &H0     '04

For i = 24 To 30
    CBW(i) = 0
Next

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
 
Dim result As Long

'1. CBW command

 
result = WriteFile _
       (WriteHandle, _
       CBW(0), _
       31, _
       NumberOfBytesWritten, _
       0)    'out

If result = 0 Then
 InquiryScSiString = 0
 Exit Function
End If

'2. Readdata stage
 
result = ReadFile _
         (ReadHandle, _
          ReadData(0), _
         CBWDataTransferLength, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in


If result = 0 Then
  InquiryScSiString = 0
 Exit Function
End If




' ============ Record function
If ScSiRecMode = 0 Then

                For i = 0 To CBWDataTransferLength - 1
                Debug.Print "k", i, Hex(ReadData(i))
                 If ReadData(i) <> InquiryString(Lun, i) Then
                  
                   InquiryScSiString = 2  ' card format capacity has problem
                   Exit Function
                 End If
                
                
                Next i

                 
 

Else



' ============ Record function
 
Select Case Lun
Case 0

                Open App.Path & "\SCSI0.txt" For Output As #4

                     For i = 0 To 35
                        Print #4, ReadData(i)
                     Next i
                 Close #4
                 
Case 1
                Open App.Path & "\SCSI1.txt" For Output As #4

                     For i = 0 To 35
                        Print #4, ReadData(i)
                     Next i
                 Close #4


Case 2

                 Open App.Path & "\SCSI2.txt" For Output As #4

                     For i = 0 To 35
                        Print #4, ReadData(i)
                     Next i
                 Close #4

Case 3
                
                  Open App.Path & "\SCSI3.txt" For Output As #4

                     For i = 0 To 35
                        Print #4, ReadData(i)
                     Next i
                 Close #4



End Select

End If

'3. CSW data
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 
If result = 0 Then
 InquiryScSiString = 0
 Exit Function
End If
 
'4. CSW status

If CSW(12) = 1 Then
     InquiryScSiString = 0
Else
     InquiryScSiString = 1
   
End If

 
End Function





Function Check_Erase_Block0(LBA As Long, Lun As Byte, CBWDataTransferLength As Long) As Byte
Dim CBW(0 To 30) As Byte
Dim NumberOfBytesWritten As Long
Dim CBWDataTransferLen(0 To 3) As Byte
  
Dim TransferLen As Long
Dim TransferLenLSB As Byte
Dim TransferLenMSB As Byte
Dim i As Integer
Dim tmpV(0 To 2) As Long
Dim opcode As Byte

Dim CSW(0 To 12) As Byte

Dim NumberOfBytesRead As Long

For i = 0 To 30
   
        CBW(i) = 0
    
Next i

For i = 0 To CBWDataTransferLength
ReadData(i) = 0
Next

Const CBWSignature_0 = &H55
Const CBWSignature_1 = &H53
Const CBWSignature_2 = &H42
Const CBWSignature_3 = &H43


Const CBWTag_0 = &H1
Const CBWTag_1 = &H2
Const CBWTag_2 = &H3
Const CBWTag_3 = &H4


'/////////////////// CBW signature

CBW(0) = CBWSignature_0
CBW(1) = CBWSignature_1
CBW(2) = CBWSignature_2
CBW(3) = CBWSignature_3

'/////////////////  CBW Tag

CBW(4) = CBWTag_0
CBW(5) = CBWTag_1
CBW(6) = CBWTag_2
CBW(7) = CBWTag_3

 

CBW(8) = &H81  '00
CBW(9) = &H0   '08
CBW(10) = &H2   '00
CBW(11) = &H0  '00

'///////////////  CBW Flag
CBW(12) = &H80                 '80

'////////////// LUN
CBW(13) = &H0
'///////////// CBD Len
CBW(14) = &H10                '0a

'////////////  UFI command

CBW(15) = &HD0
CBW(16) = &H0


CBW(17) = &HF0        '00
CBW(18) = &H70         '00
CBW(19) = &HFE         '00
CBW(20) = &H3       '40

'/////////////  Reverve
CBW(21) = &H0

'//////////// Transfer Len


CBW(22) = &H0      '00
CBW(23) = &H0       '04
CBW(24) = &HF0       '04
CBW(25) = &HD0       '04
CBW(26) = &HF3       '04
CBW(27) = &H1        '04
CBW(28) = &H0       '04
CBW(29) = &H0       '04
CBW(30) = &H0       '04



'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
 
Dim result As Long

'1. CBW command

 
result = WriteFile _
       (WriteHandle, _
       CBW(0), _
       31, _
       NumberOfBytesWritten, _
       0)    'out

If result = 0 Then
 'Read_Data = 0
 Exit Function
End If

'2. Readdata stage
 
result = ReadFile _
         (ReadHandle, _
          ReadData(0), _
         CBWDataTransferLength, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in

 
'If result = 0 Then
' Read_Data = 0
' Exit Function
'End If

'3. CSW data
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 
If result = 0 Then
 'Read_Data = 0
 Exit Function
End If
 
'4. CSW status

If CSW(12) = 1 Then
   ' Read_Data = 0
Else
    ' Read_Data = 1
   
End If

 
End Function
Function Erase_Block0(LBA As Long, Lun As Byte, CBWDataTransferLength As Long) As Byte
Dim CBW(0 To 30) As Byte
Dim NumberOfBytesWritten As Long
Dim CBWDataTransferLen(0 To 3) As Byte
  
Dim TransferLen As Long
Dim TransferLenLSB As Byte
Dim TransferLenMSB As Byte
Dim i As Integer
Dim tmpV(0 To 2) As Long
Dim opcode As Byte

Dim CSW(0 To 12) As Byte

Dim NumberOfBytesRead As Long

For i = 0 To 30
   
        CBW(i) = 0
    
Next i

For i = 0 To CBWDataTransferLength
ReadData(i) = 0
Next

Const CBWSignature_0 = &H55
Const CBWSignature_1 = &H53
Const CBWSignature_2 = &H42
Const CBWSignature_3 = &H43


Const CBWTag_0 = &H1
Const CBWTag_1 = &H2
Const CBWTag_2 = &H3
Const CBWTag_3 = &H4


'/////////////////// CBW signature

CBW(0) = CBWSignature_0
CBW(1) = CBWSignature_1
CBW(2) = CBWSignature_2
CBW(3) = CBWSignature_3

'/////////////////  CBW Tag

CBW(4) = CBWTag_0
CBW(5) = CBWTag_1
CBW(6) = CBWTag_2
CBW(7) = CBWTag_3

 

CBW(8) = &H81  '00
CBW(9) = &H0   '08
CBW(10) = &H2   '00
CBW(11) = &H0  '00

'///////////////  CBW Flag
CBW(12) = &H80                 '80

'////////////// LUN
CBW(13) = &H0
'///////////// CBD Len
CBW(14) = &H10                '0a

'////////////  UFI command

CBW(15) = &HD0
CBW(16) = &H0


CBW(17) = &HF0        '00
CBW(18) = &H60         '00
CBW(19) = &HFE         '00
CBW(20) = &H3       '40

'/////////////  Reverve
CBW(21) = &H0

'//////////// Transfer Len


CBW(22) = &H0      '00
CBW(23) = &H0       '04
CBW(24) = &HF0       '04
CBW(25) = &HD0       '04
CBW(26) = &HF3       '04
CBW(27) = &H1        '04
CBW(28) = &H0       '04
CBW(29) = &H0       '04
CBW(30) = &H0       '04



'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
 
Dim result As Long

'1. CBW command

 
result = WriteFile _
       (WriteHandle, _
       CBW(0), _
       31, _
       NumberOfBytesWritten, _
       0)    'out

If result = 0 Then
 'Read_Data = 0
 Exit Function
End If

'2. Readdata stage
 
result = ReadFile _
         (ReadHandle, _
          ReadData(0), _
         CBWDataTransferLength, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in

 
'If result = 0 Then
' Read_Data = 0
' Exit Function
'End If

'3. CSW data
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 
If result = 0 Then
' Read_Data = 0
 Exit Function
End If
 
'4. CSW status

If CSW(12) = 1 Then
   ' Read_Data = 0
Else
    ' Read_Data = 1
   
End If

 
End Function
Function Phyical_Read_Firmware(LBA As Long, Lun As Byte, CBWDataTransferLength As Long) As Byte
Dim CBW(0 To 30) As Byte
Dim NumberOfBytesWritten As Long
Dim CBWDataTransferLen(0 To 3) As Byte
  
Dim TransferLen As Long
Dim TransferLenLSB As Byte
Dim TransferLenMSB As Byte
Dim i As Integer
Dim tmpV(0 To 2) As Long
Dim opcode As Byte

Dim CSW(0 To 12) As Byte

Dim NumberOfBytesRead As Long

For i = 0 To 30
   
        CBW(i) = 0
    
Next i

For i = 0 To CBWDataTransferLength
ReadData(i) = 0
Next

Const CBWSignature_0 = &H55
Const CBWSignature_1 = &H53
Const CBWSignature_2 = &H42
Const CBWSignature_3 = &H43


Const CBWTag_0 = &H1
Const CBWTag_1 = &H2
Const CBWTag_2 = &H3
Const CBWTag_3 = &H4


'/////////////////// CBW signature

CBW(0) = CBWSignature_0
CBW(1) = CBWSignature_1
CBW(2) = CBWSignature_2
CBW(3) = CBWSignature_3

'/////////////////  CBW Tag

CBW(4) = CBWTag_0
CBW(5) = CBWTag_1
CBW(6) = CBWTag_2
CBW(7) = CBWTag_3

CBWDataTransferLen(0) = (CBWDataTransferLength Mod 256)
tmpV(0) = Int(CBWDataTransferLength / 256)
CBWDataTransferLen(1) = (tmpV(0) Mod 256)
tmpV(1) = Int(tmpV(0) / 256)
CBWDataTransferLen(2) = (tmpV(1) Mod 256)
tmpV(2) = Int((tmpV(1) / 256))
CBWDataTransferLen(3) = (tmpV(2) Mod 256)

CBW(8) = CBWDataTransferLen(0)  '00
CBW(9) = CBWDataTransferLen(1)  '08
CBW(10) = CBWDataTransferLen(2) '00
CBW(11) = CBWDataTransferLen(3) '00

'///////////////  CBW Flag
CBW(12) = &H80                 '80

'////////////// LUN
CBW(13) = Lun                    '00

'///////////// CBD Len
CBW(14) = &HA                '0a

'////////////  UFI command

CBW(15) = &H28
CBW(16) = Lun * 32
LBAByte(0) = (LBA Mod 256)
tmpV(0) = Int(LBA / 256)
LBAByte(1) = (tmpV(0) Mod 256)
tmpV(1) = Int(tmpV(0) / 256)
LBAByte(2) = (tmpV(1) Mod 256)
tmpV(2) = Int((tmpV(1) / 256))
LBAByte(3) = (tmpV(2) Mod 256)

CBW(17) = LBAByte(3)         '00
CBW(18) = LBAByte(2)         '00
CBW(19) = LBAByte(1)         '00
CBW(20) = LBAByte(0)         '40

'/////////////  Reverve
CBW(21) = 0

'//////////// Transfer Len

TransferLen = Int(CBWDataTransferLength / 512)

TransferLenLSB = (TransferLen Mod 256)
tmpV(0) = Int(TransferLen / 256)
TransferLenMSB = (tmpV(0) / 256)

CBW(22) = TransferLenMSB      '00
CBW(23) = TransferLenLSB      '04

For i = 24 To 30
    CBW(i) = 0
Next

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
 
Dim result As Long

'1. CBW command

 
result = WriteFile _
       (WriteHandle, _
       CBW(0), _
       31, _
       NumberOfBytesWritten, _
       0)    'out

If result = 0 Then
' Read_Data = 0
 Exit Function
End If

'2. Readdata stage
 
result = ReadFile _
         (ReadHandle, _
          ReadData(0), _
         CBWDataTransferLength, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in

 
'If result = 0 Then
' Read_Data = 0
' Exit Function
'End If

'3. CSW data
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 
If result = 0 Then
' Read_Data = 0
 Exit Function
End If
 
'4. CSW status

If CSW(12) = 1 Then
    'Read_Data = 0
Else
    ' Read_Data = 1
   
End If

 
End Function




Function Physical_Read_Data(LBA As Long, Lun As Byte, CBWDataTransferLength As Long) As Byte
Dim CBW(0 To 30) As Byte
Dim NumberOfBytesWritten As Long
Dim CBWDataTransferLen(0 To 3) As Byte
  
Dim TransferLen As Long
Dim TransferLenLSB As Byte
Dim TransferLenMSB As Byte
Dim i As Integer
Dim tmpV(0 To 2) As Long
Dim opcode As Byte

Dim CSW(0 To 12) As Byte

Dim NumberOfBytesRead As Long

For i = 0 To 30
   
        CBW(i) = 0
    
Next i

For i = 0 To CBWDataTransferLength
ReadData(i) = 0
Next

Const CBWSignature_0 = &H55
Const CBWSignature_1 = &H53
Const CBWSignature_2 = &H42
Const CBWSignature_3 = &H43


Const CBWTag_0 = &H1
Const CBWTag_1 = &H2
Const CBWTag_2 = &H3
Const CBWTag_3 = &H4


'/////////////////// CBW signature

CBW(0) = CBWSignature_0
CBW(1) = CBWSignature_1
CBW(2) = CBWSignature_2
CBW(3) = CBWSignature_3

'/////////////////  CBW Tag

CBW(4) = CBWTag_0
CBW(5) = CBWTag_1
CBW(6) = CBWTag_2
CBW(7) = CBWTag_3

CBWDataTransferLen(0) = (CBWDataTransferLength Mod 256)
tmpV(0) = Int(CBWDataTransferLength / 256)
CBWDataTransferLen(1) = (tmpV(0) Mod 256)
tmpV(1) = Int(tmpV(0) / 256)
CBWDataTransferLen(2) = (tmpV(1) Mod 256)
tmpV(2) = Int((tmpV(1) / 256))
CBWDataTransferLen(3) = (tmpV(2) Mod 256)

CBW(8) = CBWDataTransferLen(0)  '00
CBW(9) = CBWDataTransferLen(1)  '08
CBW(10) = CBWDataTransferLen(2) '00
CBW(11) = CBWDataTransferLen(3) '00

'///////////////  CBW Flag
CBW(12) = &H80                 '80

'////////////// LUN
CBW(13) = Lun                    '00

'///////////// CBD Len
CBW(14) = &HA                '0a

'////////////  UFI command

CBW(15) = &H9C  ' physical read command
CBW(16) = &H0
LBAByte(0) = (LBA Mod 256)
tmpV(0) = Int(LBA / 256)
LBAByte(1) = (tmpV(0) Mod 256)
tmpV(1) = Int(tmpV(0) / 256)
LBAByte(2) = (tmpV(1) Mod 256)
tmpV(2) = Int((tmpV(1) / 256))
LBAByte(3) = (tmpV(2) Mod 256)

CBW(17) = 0        '00
CBW(18) = 0        '00
CBW(19) = 0         '00
CBW(20) = 0         '40

'/////////////  Reverve
CBW(21) = 0

'//////////// Transfer Len

TransferLen = Int(CBWDataTransferLength / 512)

TransferLenLSB = (TransferLen Mod 256)
tmpV(0) = Int(TransferLen / 256)
TransferLenMSB = (tmpV(0) / 256)

CBW(22) = 0     '00
CBW(23) = 0      '04

For i = 24 To 30
    CBW(i) = 0
Next

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
 
Dim result As Long

'1. CBW command

 
result = WriteFile _
       (WriteHandle, _
       CBW(0), _
       31, _
       NumberOfBytesWritten, _
       0)    'out

If result = 0 Then
 Physical_Read_Data = 0
 Exit Function
End If

'2. Readdata stage
 
result = ReadFile _
         (ReadHandle, _
          ReadData(0), _
         CBWDataTransferLength, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in

 
'If result = 0 Then
' Read_Data = 0
' Exit Function
'End If

'3. CSW data
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 
If result = 0 Then
 Physical_Read_Data = 0
 Exit Function
End If
 
'4. CSW status

If CSW(12) = 1 Then
    Physical_Read_Data = 0
Else
     Physical_Read_Data = 1
   
End If

 
End Function

Function AU6981_Scan(ChipNo As Byte, Zone As Byte) As Byte
Dim CBW(0 To 30) As Byte
Dim NumberOfBytesWritten As Long
Dim CBWDataTransferLen(0 To 3) As Byte
  
Dim TransferLen As Long
Dim TransferLenLSB As Byte
Dim TransferLenMSB As Byte
Dim i As Integer
Dim tmpV(0 To 2) As Long
Dim opcode As Byte

Dim CSW(0 To 12) As Byte

Dim NumberOfBytesRead As Long

For i = 0 To 30
   
        CBW(i) = 0
    
Next i
 
For i = 0 To 511
ReadData(i) = 0
Next

Const CBWSignature_0 = &H55
Const CBWSignature_1 = &H53
Const CBWSignature_2 = &H42
Const CBWSignature_3 = &H43


Const CBWTag_0 = &H1
Const CBWTag_1 = &H2
Const CBWTag_2 = &H3
Const CBWTag_3 = &H4


'/////////////////// CBW signature

CBW(0) = CBWSignature_0
CBW(1) = CBWSignature_1
CBW(2) = CBWSignature_2
CBW(3) = CBWSignature_3

'/////////////////  CBW Tag

CBW(4) = CBWTag_0
CBW(5) = CBWTag_1
CBW(6) = CBWTag_2
CBW(7) = CBWTag_3

 
CBW(8) = 0   '00
CBW(9) = 2  '08
CBW(10) = 0 '00
CBW(11) = 0  '00

'///////////////  CBW Flag
CBW(12) = &H80                 '80

'////////////// LUN
CBW(13) = 0                  '00

'///////////// CBD Len
CBW(14) = &H10               '0a

'////////////  UFI command

CBW(15) = &H94  ' physical read command
CBW(16) = ChipNo
LBAByte(0) = (LBA Mod 256)
tmpV(0) = Int(LBA / 256)
LBAByte(1) = (tmpV(0) Mod 256)
tmpV(1) = Int(tmpV(0) / 256)
LBAByte(2) = (tmpV(1) Mod 256)
tmpV(2) = Int((tmpV(1) / 256))
LBAByte(3) = (tmpV(2) Mod 256)

CBW(17) = Zone        '00
CBW(18) = 0        '00
CBW(19) = 0         '00
CBW(20) = 0         '40

'/////////////  Reverve
CBW(21) = 0

'//////////// Transfer Len

 
CBW(22) = 0     '00
CBW(23) = 0      '04

For i = 24 To 30
    CBW(i) = 0
Next

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
 
Dim result As Long

'1. CBW command

 
result = WriteFile _
       (WriteHandle, _
       CBW(0), _
       31, _
       NumberOfBytesWritten, _
       0)    'out

If result = 0 Then
 AU6981_Scan = 0
 Exit Function
End If

'2. Readdata stage
 
result = ReadFile _
         (ReadHandle, _
          ReadData(0), _
         512, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in

 
 If result = 0 Then
  AU6981_Scan = 0
  Exit Function
 End If

'3. CSW data
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 
If result = 0 Then
 AU6981_Scan = 0
 Exit Function
End If
 
'4. CSW status

If CSW(12) = 1 Then
    AU6981_Scan = 0
   
    
    
    
Else
     AU6981_Scan = 1
     
      ' get bad block poisition
    BadBlockCounter = 0
    For i = 1 To 511  ' 0th data is zone index
    
    If (ReadData(i) <> &HAA) And (ReadData(i + 1) <> &H55) Then
         BadBlockCounter = BadBlockCounter + 1
         BadBlock(i - 1) = ReadData(i)
    Else
         
         Exit Function
    End If
    
    Next i
 
   
End If

 
  
   
 
End Function

Function AU6981_Fix(ChipNo As Byte, Zone As Byte, BadBlockPos As Byte) As Byte
Dim CBW(0 To 30) As Byte
Dim NumberOfBytesWritten As Long
Dim CBWDataTransferLen(0 To 3) As Byte
  
Dim TransferLen As Long
Dim TransferLenLSB As Byte
Dim TransferLenMSB As Byte
Dim i As Integer
Dim tmpV(0 To 2) As Long
Dim opcode As Byte

Dim CSW(0 To 12) As Byte

Dim NumberOfBytesRead As Long

For i = 0 To 30
   
        CBW(i) = 0
    
Next i
 
For i = 0 To 511
ReadData(i) = 0
Next

Const CBWSignature_0 = &H55
Const CBWSignature_1 = &H53
Const CBWSignature_2 = &H42
Const CBWSignature_3 = &H43


Const CBWTag_0 = &H1
Const CBWTag_1 = &H2
Const CBWTag_2 = &H3
Const CBWTag_3 = &H4


'/////////////////// CBW signature

CBW(0) = CBWSignature_0
CBW(1) = CBWSignature_1
CBW(2) = CBWSignature_2
CBW(3) = CBWSignature_3

'/////////////////  CBW Tag

CBW(4) = CBWTag_0
CBW(5) = CBWTag_1
CBW(6) = CBWTag_2
CBW(7) = CBWTag_3

 
CBW(8) = 0   '00
CBW(9) = 2  '08
CBW(10) = 0 '00
CBW(11) = 0  '00

'///////////////  CBW Flag
CBW(12) = &H80                 '80

'////////////// LUN
CBW(13) = 0                  '00

'///////////// CBD Len
CBW(14) = &H10               '0a

'////////////  UFI command

CBW(15) = &H98  ' physical read command
CBW(16) = ChipNo
 

CBW(17) = Zone        '00
CBW(18) = BadBlockPos      '00
CBW(19) = &H15        '00
CBW(20) = 0         '40

'/////////////  Reverve
CBW(21) = 0

'//////////// Transfer Len

 
CBW(22) = 0     '00
CBW(23) = 1      '04

For i = 24 To 30
    CBW(i) = 0
Next

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
 
Dim result As Long

'1. CBW command

 
result = WriteFile _
       (WriteHandle, _
       CBW(0), _
       31, _
       NumberOfBytesWritten, _
       0)    'out

If result = 0 Then
 AU6981_Fix = 0
 Exit Function
End If

'2. Readdata stage
 
result = ReadFile _
         (ReadHandle, _
          ReadData(0), _
         512, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in

 
 If result = 0 Then
  AU6981_Fix = 0
  Exit Function
 End If

'3. CSW data
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 
If result = 0 Then
 AU6981_Fix = 0
 Exit Function
End If
 
'4. CSW status

If CSW(12) = 1 Then
    AU6981_Fix = 0
  
Else
     AU6981_Fix = 1
     
   
   
End If

 
  
   
 
End Function





Function ClosePipe6331() As Integer
On Error Resume Next
 
CloseHandle (WriteHandle6331)


End Function


Function AU6375EndLess(Lun As Byte) As Byte

On Error GoTo CancelHandle
Dim CBW(0 To 30) As Byte
Dim CSW(0 To 12) As Byte
Dim i As Integer
Dim NumberOfBytesWritten As Long
Dim NumberOfBytesRead As Long
Dim result As Long

     For i = 0 To 30
    
        CBW(i) = 0
    
    Next i

CBW(0) = &H55 'signature
CBW(1) = &H53
CBW(2) = &H42
'CBW(3) = &H43  ' orgioinal tag
CBW(3) = &H43 ' for test unit speed


CBW(4) = &H8   'package ID
CBW(5) = &HB0
CBW(6) = &H32
CBW(7) = &H84


CBW(8) = &H18
CBW(9) = &H0
CBW(10) = &H0
CBW(11) = &H0


CBW(12) = &H80  '    CBW FLAG 0000
 
CBW(13) = &H1
CBW(14) = &HA
CBW(15) = &H76

CBW(16) = &H4
CBW(17) = &H30
CBW(18) = &H35
CBW(19) = &H38

CBW(20) = &H46
CBW(21) = &H1
 
 


'1. CBW output

AU6375EndLess = 0
If ChipName = "AU6375AS" Then

        result = WriteFile _
               (WriteHandle, _
               CBW(0), _
               31, _
               NumberOfBytesWritten, _
               0)
        
        If result = 0 Then
             AU6375EndLess = 0
            Exit Function
        End If

End If


'2. CSW input
 OldTimer = Timer
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
Debug.Print "AU6375EndLess"
For i = 0 To 12
Debug.Print i, CSW(i)
Next i

If result = 0 Then
    AU6375EndLess = 0
   Exit Function
End If

 
'3 CSW Status
If CSW(12) = 1 Then
    AU6375EndLess = 0
    
    Else
    AU6375EndLess = 1
    End If

Exit Function

CancelHandle:

CancelIo (ReadHandle)



End Function

Function AU6375TestMode(Lun As Byte) As Byte
Dim CBW(0 To 30) As Byte
Dim CSW(0 To 12) As Byte
Dim i As Integer
Dim NumberOfBytesWritten As Long
Dim NumberOfBytesRead As Long
Dim result As Long

     For i = 0 To 30
    
        CBW(i) = 0
    
    Next i

CBW(0) = &H55 'signature
CBW(1) = &H53
CBW(2) = &H42
'CBW(3) = &H43  ' orgioinal tag
CBW(3) = &H43 ' for test unit speed


CBW(4) = &H8   'package ID
CBW(5) = &HB0
CBW(6) = &H32
CBW(7) = &H84


CBW(8) = &H0
CBW(9) = &H0
CBW(10) = &H0
CBW(11) = &H0


CBW(12) = &H0  '    CBW FLAG 0000
 
CBW(13) = &H1
CBW(14) = &HA
CBW(15) = &HC7

CBW(16) = &H4
CBW(17) = &H30
CBW(18) = &H35
CBW(19) = &H38

CBW(20) = &H46
CBW(21) = &H1
 
 


'1. CBW output

AU6375TestMode = 0

result = WriteFile _
       (WriteHandle, _
       CBW(0), _
       31, _
       NumberOfBytesWritten, _
       0)

If result = 0 Then
     AU6375TestMode = 0
    Exit Function
End If


'2. CSW input
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
Debug.Print "AU6375TestMode"
For i = 0 To 12
Debug.Print i, CSW(i)
Next i

If result = 0 Then
    AU6375TestMode = 0
   Exit Function
End If

 
'3 CSW Status
If CSW(4) = &H8 And CSW(5) = &HB0 And CSW(6) = &H32 And CSW(7) = &H84 Then
    AU6375TestMode = 1
    
    Else
    AU6375TestMode = 0
    End If

End Function




Function CBWTest_New_8_Sector(Lun As Byte, PreSlotStatus As Byte, FailPosition As Integer, Times As Integer) As Byte
Dim i As Integer
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long
Dim j As Integer

 CBWDataTransferLength = 1024 * 4 ' 8 sector

   
    If PreSlotStatus <> 1 Then
        CBWTest_New_8_Sector = 4
        Exit Function
    End If
    '========================================
   
    CBWTest_New_8_Sector = 2
   
    '========================================
    
    
     If OpenPipe = 0 Then
       CBWTest_New_8_Sector = 2   ' Write fail
       Exit Function
     End If
  
    '====================================
     TmpInteger = TestUnitSpeed(Lun)
    
    If TmpInteger = 0 Then
        
       CBWTest_New_8_Sector = 2   ' usb 2.0 high speed fail
       UsbSpeedTestResult = 2
       Exit Function
    End If
    TmpInteger = 0
    
    TmpInteger = TestUnitReady(Lun)
     If TmpInteger = 0 Then
         TmpInteger = RequestSense(Lun)
        
         If TmpInteger = 0 Then
        
            CBWTest_New_8_Sector = 2  'Write fail
            Exit Function
         End If
        
     End If
  
  For j = 1 To Times
   
   
        If LBA > 25 * 1024 Then
            LBA = 0
        End If
    
       ' For i = 0 To CBWDataTransferLength - 1
        
       '      ReadData(i) = 0
    
       ' Next

      
        TmpInteger = Write_Data(LBA, Lun, CBWDataTransferLength)
         
        If TmpInteger = 0 Then
            CBWTest_New_8_Sector = 2  'write fail
            FailPosition = j
            Exit Function
        End If
    
        TmpInteger = Read_Data(LBA, Lun, CBWDataTransferLength)
         
        If TmpInteger = 0 Then
            CBWTest_New_8_Sector = 3    'Read fail
            FailPosition = j
            Exit Function
        End If
     
        For i = 0 To CBWDataTransferLength - 1
        
            If ReadData(i) <> Pattern(i) Then
              CBWTest_New_8_Sector = 3    'Read fail
              FailPosition = j
              Exit Function
            End If
        
        Next
        
        CBWTest_New_8_Sector = 1
           
        LBA = LBA + 1
  Next
    
    End Function





Function AU6375_GetDevice(Lun As Byte, PreSlotStatus As Byte, Vid_PID As String) As Byte
Dim i As Integer
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long

   CBWDataTransferLength = 1024
 
'   For i = 0 To CBWDataTransferLength - 1
    
'         ReadData(i) = 0

'   Next

    If PreSlotStatus <> 1 Then
        AU6375_GetDevice = 4
        Exit Function
    End If
    '========================================
   
    AU6375_GetDevice = 0
    If LBA > 25 * 1024 Then
        LBA = 0
    End If
    '========================================
     TmpString = ""
    If ReaderExist = 0 Then
        Do
            DoEvents
            Call MsecDelay(0.1)
            TimerCounter = TimerCounter + 1
            TmpString = GetDeviceName(Vid_PID)
        Loop While TmpString = "" And TimerCounter < 10
    End If
    '=======================================
    If ReaderExist = 0 And TmpString <> "" Then
      ReaderExist = 1
    End If
    '=======================================
    If ReaderExist = 0 And TmpString = "" Then
      AU6375_GetDevice = 0   ' no readerExist
      ReaderExist = 0
      Exit Function
    End If
    '=======================================
    If OpenPipe = 0 Then
      AU6375_GetDevice = 5   ' Write fail
      Exit Function
    End If
 
    '======================================
    
     ' for unitSpeed
    
    TmpInteger = TestUnitSpeed(Lun)
    
    If TmpInteger = 0 Then
        
       AU6375_GetDevice = 2   ' usb 2.0 high speed fail
       UsbSpeedTestResult = 2
       Exit Function
    End If
    
    
    '==== Rec mode
   '   TmpInteger = InquiryScSiString(0, 36, 1)
   '   TmpInteger = InquiryScSiString(1, 36, 1)
   '   TmpInteger = InquiryScSiString(2, 36, 1)
   '   TmpInteger = InquiryScSiString(3, 36, 1)
      
      
   If ChipName = "AU6375AS" Then
   TmpInteger = InquiryScSiString(0, 36, 0)
     If TmpInteger = 0 Then
        
        AU6375_GetDevice = 2   ' usb 2.0 high speed fail
        
       Exit Function
     End If
    
    
     TmpInteger = InquiryScSiString(1, 36, 0)
     If TmpInteger = 0 Then
        
        AU6375_GetDevice = 2   ' usb 2.0 high speed fail
        
       Exit Function
     End If
     
      TmpInteger = InquiryScSiString(2, 36, 0)
     If TmpInteger = 0 Then
        
        AU6375_GetDevice = 2   ' usb 2.0 high speed fail
        
       Exit Function
     End If
     
     
     TmpInteger = InquiryScSiString(3, 36, 0)
     If TmpInteger = 0 Then
        
        AU6375_GetDevice = 2   ' usb 2.0 high speed fail
        
       Exit Function
     End If
     
   End If
     
     
    AU6375_GetDevice = 1
        
    
    End Function




Function CBWTest_NewAU9331(Lun As Byte, PreSlotStatus As Byte, Vid_PID As String) As Byte
Dim i As Integer
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long

   CBWDataTransferLength = 1024
 
'   For i = 0 To CBWDataTransferLength - 1
    
'         ReadData(i) = 0

'   Next

    If PreSlotStatus <> 1 Then
        CBWTest_NewAU9331 = 4
        Exit Function
    End If
    '========================================
   
    CBWTest_NewAU9331 = 0
    If LBA > 25 * 1024 Then
        LBA = 0
    End If
    '========================================
     TmpString = ""
    If ReaderExist = 0 Then
        Do
            DoEvents
            Call MsecDelay(0.1)
            TimerCounter = TimerCounter + 1
            TmpString = GetDeviceName(Vid_PID)
        Loop While TmpString = "" And TimerCounter < 10
    End If
    '=======================================
    If ReaderExist = 0 And TmpString <> "" Then
      ReaderExist = 1
    End If
    '=======================================
    If ReaderExist = 0 And TmpString = "" Then
      CBWTest_NewAU9331 = 0   ' no readerExist
      ReaderExist = 0
      Exit Function
    End If
    '=======================================
    If OpenPipe = 0 Then
      CBWTest_NewAU9331 = 2   ' Write fail
      Exit Function
    End If
 
    '======================================
    
     ' for unitSpeed
    
     TmpInteger = TestUnitSpeed(Lun)
    
     If TmpInteger = 0 Then
        
        CBWTest_NewAU9331 = 2   ' usb 2.0 high speed fail
        UsbSpeedTestResult = 2
        Exit Function
     End If
    
    
    
    TmpInteger = TestUnitReady(Lun)
    If TmpInteger = 0 Then
        TmpInteger = RequestSense(Lun)
        
        If TmpInteger = 0 Then
        
           CBWTest_NewAU9331 = 2  'Write fail
           Exit Function
        End If
        
    End If
    '======================================
  '  If ChipName = "AU6371" Or ChipName = "AU6371S3" Then
  '      TmpInteger = Read_Data1(LBA, Lun, CBWDataTransferLength)
  '  End If
    TmpInteger = 0
    TmpInteger = Read_DataAU9331(LBA, Lun, CBWDataTransferLength)
     
    
    CBWTest_NewAU9331 = TmpInteger
        
    
    End Function
Function CBWTest_New_Physical_Read(Lun As Byte, PreSlotStatus As Byte, Vid_PID As String) As Byte
Dim i As Integer
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long

   CBWDataTransferLength = 512
 
'   For i = 0 To CBWDataTransferLength - 1
    
'         ReadData(i) = 0

'   Next

    If PreSlotStatus <> 1 Then
        CBWTest_New_Physical_Read = 4
        Exit Function
    End If
    '========================================
   
    CBWTest_New_Physical_Read = 0
    If LBA > 25 * 1024 Then
        LBA = 0
    End If
    '========================================
     TmpString = ""
    If ReaderExist = 0 Then
        Do
            DoEvents
            Call MsecDelay(0.1)
            TimerCounter = TimerCounter + 1
            TmpString = GetDeviceName(Vid_PID)
        Loop While TmpString = "" And TimerCounter < 10
    End If
    '=======================================
    If ReaderExist = 0 And TmpString <> "" Then
      ReaderExist = 1
    End If
    '=======================================
    If ReaderExist = 0 And TmpString = "" Then
      CBWTest_New_Physical_Read = 0   ' no readerExist
      ReaderExist = 0
      Exit Function
    End If
    '=======================================
    If OpenPipe = 0 Then
      CBWTest_New_Physical_Read = 2   ' Write fail
      Exit Function
    End If
 
    '======================================
    
     ' for unitSpeed
    
     TmpInteger = TestUnitSpeed(Lun)
    
     If TmpInteger = 0 Then
        
        CBWTest_New_Physical_Read = 2  ' usb 2.0 high speed fail
        UsbSpeedTestResult = 2
        Exit Function
     End If
    
    
    
    TmpInteger = TestUnitReady(Lun)
    If TmpInteger = 0 Then
        TmpInteger = RequestSense(Lun)
        
        If TmpInteger = 0 Then
        
           CBWTest_New_Physical_Read = 2  'Write fail
           Exit Function
        End If
        
    End If
    '======================================
  '  If ChipName = "AU6371" Or ChipName = "AU6371S3" Then
  '      TmpInteger = Read_Data1(LBA, Lun, CBWDataTransferLength)
  '  End If
    
    
    
    TmpInteger = Physical_Read_Data(LBA, 0, 512)
    If TmpInteger = 0 Then
        CBWTest_New_Physical_Read = 3    'Read fail
        Exit Function
    End If
     
     For i = 0 To CBWDataTransferLength - 1
    
   
   
       If ReadData(i) <> AU6981Pattern(i) Then
          CBWTest_New_Physical_Read = 3    'Read fail
          Exit Function
        End If
    
    Next
     
     
  
    CBWTest_New_Physical_Read = 1
        
     
    End Function



Function CBWTransfer_8_Block() As Byte
On Error Resume Next
Dim i As Integer
Dim TransferLen As Long
Dim TransferLenLSB As Byte
Dim TransferLenMSB As Byte
Dim rv As Byte
Dim tmpV(0 To 2) As Long
'to test reader by vary different address, because the memory card are demaged very serious

'////////////// Define const

Const CBWSignature_0 = &H55
Const CBWSignature_1 = &H53
Const CBWSignature_2 = &H42
Const CBWSignature_3 = &H43


Const CBWTag_0 = &H1
Const CBWTag_1 = &H2
Const CBWTag_2 = &H3
Const CBWTag_3 = &H4

Dim CBWDataTransferLength As Long
CBWDataTransferLength = 4096

'LBA = &H20
'OPCode = &H2A
'Lun = 0
'CDBLen = 10
'CBWFlag = &H0


rv = 0
For i = 0 To CBWDataTransferLength - 1
TransData(i) = Pattern(i)
'Debug.Print TransData(i)
Next i

 

'///////////// parameter list


'/////////////////// CBW signature

CBW(0) = CBWSignature_0
CBW(1) = CBWSignature_1
CBW(2) = CBWSignature_2
CBW(3) = CBWSignature_3

'/////////////////  CBW Tag

CBW(4) = CBWTag_0
CBW(5) = CBWTag_1
CBW(6) = CBWTag_2
CBW(7) = CBWTag_3

'////////////////  CBW Data Transfer


CBWDataTransferLen(0) = (CBWDataTransferLength Mod 256)
tmpV(0) = Int(CBWDataTransferLength / 256)
CBWDataTransferLen(1) = (tmpV(0) Mod 256)
tmpV(1) = Int(tmpV(0) / 256)
CBWDataTransferLen(2) = (tmpV(1) Mod 256)
tmpV(2) = Int((tmpV(1) / 256))
CBWDataTransferLen(3) = (tmpV(2) Mod 256)

CBW(8) = CBWDataTransferLen(0)  '00
CBW(9) = CBWDataTransferLen(1)  '08
CBW(10) = CBWDataTransferLen(2) '00
CBW(11) = CBWDataTransferLen(3) '00

'///////////////  CBW Flag
CBW(12) = CBWFlag                '80

'////////////// LUN
CBW(13) = Lun                    '00

'///////////// CBD Len
CBW(14) = CDBLen                 '0a

'////////////  UFI command

CBW(15) = opcode
CBW(16) = Lun * 32

'///////////// LBA


LBAByte(0) = (LBA Mod 256)
tmpV(0) = Int(LBA / 256)
LBAByte(1) = (tmpV(0) Mod 256)
tmpV(1) = Int(tmpV(0) / 256)
LBAByte(2) = (tmpV(1) Mod 256)
tmpV(2) = Int((tmpV(1) / 256))
LBAByte(3) = (tmpV(2) Mod 256)

CBW(17) = LBAByte(3)         '00
CBW(18) = LBAByte(2)         '00
CBW(19) = LBAByte(1)         '00
CBW(20) = LBAByte(0)         '40

'/////////////  Reverve
CBW(21) = 0

'//////////// Transfer Len

TransferLen = Int(CBWDataTransferLength / 512)

TransferLenLSB = (TransferLen Mod 256)
tmpV(0) = Int(TransferLen / 256)
TransferLenMSB = (tmpV(0) / 256)

CBW(22) = TransferLenMSB      '00
CBW(23) = TransferLenLSB      '04

For i = 24 To 30
    CBW(i) = 0
Next

rv = ReaderTester2(CBW(0), TransData(0), CSW(0))


If rv = 0 Or CSW(12) = 1 Then
    CBWTransfer_8_Block = 0
Else
   
 CBWTransfer_8_Block = 1
   For i = 0 To CBWDataTransferLength - 1
   
        If TransData(i) <> Pattern(i) Then
        
         CBWTransfer_8_Block = 0
         Exit Function
        End If
   ' Debug.Print TransData(i)
   Next i

   
   

   
End If

End Function

Function CBWTransfer() As Byte
On Error Resume Next
Dim i As Integer
Dim TransferLen As Long
Dim TransferLenLSB As Byte
Dim TransferLenMSB As Byte
Dim rv As Byte
Dim tmpV(0 To 2) As Long
'to test reader by vary different address, because the memory card are demaged very serious

'////////////// Define const

Const CBWSignature_0 = &H55
Const CBWSignature_1 = &H53
Const CBWSignature_2 = &H42
Const CBWSignature_3 = &H43


Const CBWTag_0 = &H1
Const CBWTag_1 = &H2
Const CBWTag_2 = &H3
Const CBWTag_3 = &H4

Dim CBWDataTransferLength As Long
CBWDataTransferLength = 1024

'LBA = &H20
'OPCode = &H2A
'Lun = 0
'CDBLen = 10
'CBWFlag = &H0


rv = 0
For i = 0 To CBWDataTransferLength - 1
TransData(i) = Pattern(i)
'Debug.Print TransData(i)
Next i



'///////////// parameter list


'/////////////////// CBW signature

CBW(0) = CBWSignature_0
CBW(1) = CBWSignature_1
CBW(2) = CBWSignature_2
CBW(3) = CBWSignature_3

'/////////////////  CBW Tag

CBW(4) = CBWTag_0
CBW(5) = CBWTag_1
CBW(6) = CBWTag_2
CBW(7) = CBWTag_3

'////////////////  CBW Data Transfer


CBWDataTransferLen(0) = (CBWDataTransferLength Mod 256)
tmpV(0) = Int(CBWDataTransferLength / 256)
CBWDataTransferLen(1) = (tmpV(0) Mod 256)
tmpV(1) = Int(tmpV(0) / 256)
CBWDataTransferLen(2) = (tmpV(1) Mod 256)
tmpV(2) = Int((tmpV(1) / 256))
CBWDataTransferLen(3) = (tmpV(2) Mod 256)

CBW(8) = CBWDataTransferLen(0)  '00
CBW(9) = CBWDataTransferLen(1)  '08
CBW(10) = CBWDataTransferLen(2) '00
CBW(11) = CBWDataTransferLen(3) '00

'///////////////  CBW Flag
CBW(12) = CBWFlag                '80

'////////////// LUN
CBW(13) = Lun                    '00

'///////////// CBD Len
CBW(14) = CDBLen                 '0a

'////////////  UFI command

CBW(15) = opcode
CBW(16) = Lun * 32

'///////////// LBA


LBAByte(0) = (LBA Mod 256)
tmpV(0) = Int(LBA / 256)
LBAByte(1) = (tmpV(0) Mod 256)
tmpV(1) = Int(tmpV(0) / 256)
LBAByte(2) = (tmpV(1) Mod 256)
tmpV(2) = Int((tmpV(1) / 256))
LBAByte(3) = (tmpV(2) Mod 256)

CBW(17) = LBAByte(3)         '00
CBW(18) = LBAByte(2)         '00
CBW(19) = LBAByte(1)         '00
CBW(20) = LBAByte(0)         '40

'/////////////  Reverve
CBW(21) = 0

'//////////// Transfer Len

TransferLen = Int(CBWDataTransferLength / 512)

TransferLenLSB = (TransferLen Mod 256)
tmpV(0) = Int(TransferLen / 256)
TransferLenMSB = (tmpV(0) / 256)

CBW(22) = TransferLenMSB      '00
CBW(23) = TransferLenLSB      '04

For i = 24 To 30
    CBW(i) = 0
Next

CBWTransfer = 0
rv = ReaderTester(CBW(0), TransData(0), CSW(0))


If rv = 0 Then       ' unknow device
    CBWTransfer = 0
    Exit Function
End If

If rv = 2 Or CSW(12) = 1 Then   ' fail
    CBWTransfer = 2
Else
   
CBWTransfer = 1
   For i = 0 To CBWDataTransferLength - 1
   
        If TransData(i) <> Pattern(i) Then
        
         CBWTransfer = 2  'fail
         Exit Function
        End If
   ' Debug.Print TransData(i)
   Next i
  
End If

End Function
Function CBWTest_8_Block(LunNo As Byte, PreSlotStatus As Byte, FailPosition As Integer, Times As Integer) As Byte
Dim i As Integer
Dim WriteTest As Integer
Dim ReadTest As Integer
Dim TmpLBA As Integer

If PreSlotStatus <> 1 Then
    CBWTest_8_Block = 4
    Exit Function
End If
 
WriteTest = 1
ReadTest = 1
 
FailPosition = 0
 
If LBA > 25 * 1024 Then
LBA = 0
End If

'// same as test unit ready
opcode = &H28
Lun = LunNo
CDBLen = 10
CBWFlag = &H80


'CBWTransfer   ' 2 block transfer

If CBWTransfer = 0 Then
  CBWTest_8_Block = 0
 Exit Function
End If



TmpLBA = LBA
'/// Write test
If WriteTest = 1 Then

 For i = 1 To Times
   
    opcode = &H2A
    CDBLen = 10
    CBWFlag = &H0
       If CBWTransfer_8_Block <> 0 Then
       
          CBWTest_8_Block = 2
          FailPosition = i
          Exit Function
        End If
     LBA = LBA + 8
   Next i
     
End If
 
LBA = TmpLBA
If ReadTest = 1 Then

  For i = 1 To Times
    opcode = &H28
    CDBLen = 10
    CBWFlag = &H80
    
          If CBWTransfer_8_Block <> 0 Then
             FailPosition = i
            CBWTest_8_Block = 3
          Exit Function
        End If
        LBA = LBA + 8
   Next i
   
 End If
    

CBWTest_8_Block = 1
    

End Function
Function CBWTest(LunNo As Byte, PreSlotStatus As Byte) As Byte
Dim i As Integer
Dim WriteTest As Integer

CBWTest = 0
If PreSlotStatus <> 1 Then
    CBWTest = 4
    Exit Function
End If


If LBA > 25 * 1024 Then
LBA = 0
End If

' set read command
opcode = &H28
Lun = LunNo
CDBLen = 10
CBWFlag = &H80

WriteTest = 1
 

If CBWTransfer = 0 Then
  CBWTest = 0
 Exit Function
End If


'/// Write test
If WriteTest = 1 Then

   'LBA = &H0 + i * 6
    opcode = &H2A
   'Lun = i
    CDBLen = 10
    CBWFlag = &H0
    
    If CBWTransfer <> 1 Then
       'Debug.Print "fail1", Hex(OPCode), Hex(LBA)
        
       If CBWTransfer <> 1 Then
         'Debug.Print "fail2", Hex(OPCode), Hex(LBA)
          CBWTest = 2
          Exit Function
        End If
      
     End If
End If
'/// read test
'LBA = &H0 + i * 6
opcode = &H28
'Lun = i
CDBLen = 10
CBWFlag = &H80
    
    If CBWTransfer <> 1 Then
    'Debug.Print "fail1", Hex(OPCode), Hex(LBA)
    
          If CBWTransfer <> 1 Then
            'Debug.Print "fail2", Hex(OPCode), Hex(LBA)
             CBWTest = 3
             Exit Function
          End If
     
  
    End If
    
   CBWTest = 1
    

End Function

Function ErrCode(ByRef v1 As Byte) As String

Select Case v1

Case &H1A
      ErrCode = "FAIL_LBA_BASE"
      
Case &H19
      ErrCode = "FAIL_READ_ORIGINAL_CAPACITY"
      
Case &H18
      ErrCode = "FAIL_SCSI_READTOC"
      
Case &H17
      ErrCode = "FAIL_SCSI_PLAYAUD10"
      
Case &H16
      ErrCode = "FAIL_LOW_LEVEL_FORMAT"
      
Case &H15
      ErrCode = "FAIL_BAD_BLOCK"
      
Case &H14
      ErrCode = "FAIL_WRITE_PROTECT"
      
Case &H13
      ErrCode = "FAIL_ANSWER_PASSWORD"
      
Case &H12
      ErrCode = "FAIL_MAP"
      
Case &H11
      ErrCode = "FAIL_RESERVED_COMPARE"
      
Case &H10
      ErrCode = "FAIL_RESERVED_READ"
      
Case &HF
      ErrCode = "FAIL_RESERVED_WRITE"
      
Case &HE
      ErrCode = "FAIL_CROSS_COMPARE"
      
Case &HD
      ErrCode = "FAIL_CROSS_READ"
      
Case &HC
      ErrCode = "FAIL_CROSS_WRITE"
      
Case &HB
      ErrCode = "FAIL_GET_CHIP_ID"
      
Case &HA
      ErrCode = "FAIL_GENERAL_COMPARE"
      
Case &H9
      ErrCode = "FAIL_GENERAL_READ"
      
Case &H8
      ErrCode = "FAIL_GENERAL_WRITE"
      
Case &H7
      ErrCode = "FAIL_PID_VID_COMAPRE"
      
      
Case &H6
      ErrCode = "FAIL_PID_VID_WRITE"
      
      
Case &H5
      ErrCode = "FAIL_PID_VID_READ"
      
      
Case &H4
      ErrCode = "FAIL_EXECUTE_TIMEOUT"
      
      
Case &H3
      ErrCode = "FAIL_HANDLE_INVALID"
      
      
Case &H2
      ErrCode = "FAIL_DEVICE_INVALID"
      
      
Case &H1
      ErrCode = "FAIL_FLASH_INVISIBLE"
      
      
Case &H0
      ErrCode = "PASS"
      
  

End Select


End Function



Private Sub Command2_Click()
Dim vv As Single
Cls
For vv = 3.3 To 3# Step -0.01
Print vv
Next vv
Exit Sub
Shell "shutdown -s -f -t 0"
End Sub



Private Sub Check1_Click()

End Sub

Private Sub cmdPlay_Click()
Call SetVol
  If cmdPlay.Caption = "Play" Then
        mmcAudio.FileName = "D:\Documents and Settings\Administrator\桌面\New_host_PCI7248_2_960821\MP3 Pattern\test music4.mp3"
        mmcAudio.Command = "Open"
        mmcAudio.Command = "Play"
      
        cmdPlay.Caption = "Stop"
    Else
        mmcAudio.Command = "Stop"
        mmcAudio.Command = "Close"
    End If
End Sub

Private Sub Command4_Click()
End
End Sub



Private Sub Command8_Click()

'TableShow.Show
End Sub

 

Private Sub Form_Load()
Dim i As Long
Dim tmp0 As Integer
Dim tmp1 As Integer
Dim tmp2 As Integer
Dim tmp3 As Integer
Dim tmp4 As Integer
Dim k As Long

Dim ProgramName() As String
Dim U As Integer, j As Integer
Dim MyExeDate() As String
Dim ShowMyDate() As String
Dim temp As String
Dim intC As Integer
Dim strName As String

'    '取得Tester檔案名
'    For intC = Len(App.Path & "\" & App.EXEName) To 1 Step -1
'        If Mid$(App.Path & "\" & App.EXEName, intC, 1) = "\" Then
'            strName = Right(App.Path & "\" & App.EXEName, Len(App.Path & "\" & App.EXEName) - intC)
'            Exit For
'        End If
'    Next intC
'
'    '移除檔案名後的“.EXE”
'    strName = Replace(strName, ".exe", "", , , vbTextCompare)
'
'    '刪除舊的Tester與MPTester
'    Call DeleteValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", strName)
'    Call DeleteValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "MPTester")
'
'    '將檔案寫入註冊表的啟動
'    Call savestring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", strName, App.Path & "\" & App.EXEName)

    ' Get folder name for program name
    
    ' prevent multi open
    If App.EXEName <> "Tester" Or CheckMe(Me) Then
        End
    End If
    
    '20130927 for AU9540 memory leak
   
    ProgramName = Split(App.Path, "\")
    U = UBound(ProgramName)
    For i = 0 To U
        If Left(ProgramName(U - 1), 3) = "New" Then
            'Tester.Caption = "ALCOR TESTER 2012" & Right(ProgramName(U - 1), 4)
            Version_Label.Caption = Val(Mid(ProgramName(U - 1), 20, 3)) + 1911 & Right(ProgramName(U - 1), 4)
            Exit For
        Else
            MyExeDate = Split(FileDateTime(App.EXEName & ".exe"), " ")
            ShowMyDate = Split(MyExeDate(0), "/")
            For j = LBound(ShowMyDate) To UBound(ShowMyDate)
                If Len(ShowMyDate(j)) < 2 Then
                    ShowMyDate(j) = "0" & ShowMyDate(j)
                End If
                temp = temp & ShowMyDate(j)
            Next
            'Tester.Label1 = "ALCOR TESTER " & temp
            Version_Label.Caption = "ALCOR TESTER"
            Version_Label.Caption = temp
            Exit For
        End If
    Next

    Call TestNameSub

i = 0
MSComm1.CommPort = 1

MSComm1.Settings = "9600,N,8,1"
MSComm1.PortOpen = True

MSComm1.InBufferCount = 0
MSComm1.InputLen = 0
'MSComm1.DTREnable = True
'MSComm1.RTSEnable = True

MSComm2.CommPort = 2
MSComm2.Settings = "9600,N,8,1"
'MSComm2.PortOpen = True
Call ReadTable  '------------ for AU6375 test

Open App.Path & "\Pattern.txt" For Input As #2

    Do While Not EOF(2)
    
    
        Input #2, Pattern(i)
        For k = 0 To 15
        'Debug.Print k
          Pattern_64k(k * 4096 + i) = Pattern(i)
        Next k
         
        i = i + 1
    Loop

Close #2
 
'================== 64k ,128 sector
 i = 0
Open App.Path & "\AU6982_f1_s2.txt" For Input As #2

    Do While Not EOF(2)
    
         Pattern_AU6377(i) = &H5A
        Input #2, Pattern_AU6982(i)
        
         
        i = i + 1
    Loop

Close #2





'================== 64k ,128 sector
 i = 0
Open App.Path & "\AU6375_ram_unstable.txt" For Input As #2

    Do While Not EOF(2)
    
         Pattern_AU6375(i) = &HFF
        Input #2, Pattern_AU6375(i)
        
         
        i = i + 1
    Loop

Close #2
 
 
 i = 0
 Open App.Path & "\AU6981Pattern.txt" For Input As #2

    Do While Not EOF(2)
    
    
        Input #2, AU6981Pattern(i)
         
        i = i + 1
    Loop

Close #2
 
'====================

 i = 0
 Open App.Path & "\AU6371Fail.txt" For Input As #2

    Do While Not EOF(2) And i <= 65535
    
    
        Input #2, AU6371Pattern(i)
         
        i = i + 1
    Loop

Close #2

 
 
 
'==================== AU3130A
 
Open App.Path & "\MP3_A.txt" For Input As #3
Open App.Path & "\MP31_A.txt" For Input As #4
Open App.Path & "\MP32_A.txt" For Input As #5
Open App.Path & "\MP33_A.txt" For Input As #6
Open App.Path & "\MP34_A.txt" For Input As #7




   For i = 0 To 99
    
        Input #3, tmp0
        Input #4, tmp1
        Input #5, tmp2
        Input #6, tmp3
        Input #7, tmp4
        MP3Data_A(i) = CInt((CSng(tmp0) + CSng(tmp1) + CSng(tmp2) + CSng(tmp3) + CSng(tmp4)) * 0.2)
         
   Next i

Close #3
Close #4
Close #5
Close #6
Close #7


'======================= Au3130B43 1st  mode ===================================
Open App.Path & "\MP3_B.txt" For Input As #3
Open App.Path & "\MP31_B.txt" For Input As #4
Open App.Path & "\MP32_B.txt" For Input As #5
Open App.Path & "\MP33_B.txt" For Input As #6
Open App.Path & "\MP34_B.txt" For Input As #7

   For i = 0 To 99
    
        Input #3, tmp0
        Input #4, tmp1
        Input #5, tmp2
        Input #6, tmp3
        Input #7, tmp4
        MP3Data_B(i) = CInt((CSng(tmp0) + CSng(tmp1) + CSng(tmp2) + CSng(tmp3) + CSng(tmp4)) * 0.2)
         
   Next i

Close #3
Close #4
Close #5
Close #6
Close #7

'============================ AU3130B43 2nd mode ============================
Open App.Path & "\MP3_B1.txt" For Input As #3
Open App.Path & "\MP31_B1.txt" For Input As #4
Open App.Path & "\MP32_B1.txt" For Input As #5
Open App.Path & "\MP33_B1.txt" For Input As #6
Open App.Path & "\MP34_B1.txt" For Input As #7




   For i = 0 To 99
    
        Input #3, tmp0
        Input #4, tmp1
        Input #5, tmp2
        Input #6, tmp3
        Input #7, tmp4
        MP3Data_B1(i) = CInt((CSng(tmp0) + CSng(tmp1) + CSng(tmp2) + CSng(tmp3) + CSng(tmp4)) * 0.2)
         
   Next i

Close #3
Close #4
Close #5
Close #6
Close #7
 
'======================== AU3130B43 3rd mode ====================

Open App.Path & "\MP3_B2.txt" For Input As #3
Open App.Path & "\MP31_B2.txt" For Input As #4
Open App.Path & "\MP32_B2.txt" For Input As #5
Open App.Path & "\MP33_B2.txt" For Input As #6
Open App.Path & "\MP34_B2.txt" For Input As #7




   For i = 0 To 99
    
        Input #3, tmp0
        Input #4, tmp1
        Input #5, tmp2
        Input #6, tmp3
        Input #7, tmp4
        MP3Data_B2(i) = CInt((CSng(tmp0) + CSng(tmp1) + CSng(tmp2) + CSng(tmp3) + CSng(tmp4)) * 0.2)
         
   Next i

Close #3
Close #4
Close #5
Close #6
Close #7

'====================================================================

Open App.Path & "\MP3_C.txt" For Input As #3
Open App.Path & "\MP31_C.txt" For Input As #4
Open App.Path & "\MP32_C.txt" For Input As #5
Open App.Path & "\MP33_C.txt" For Input As #6
Open App.Path & "\MP34_C.txt" For Input As #7




   For i = 0 To 99
    
        Input #3, tmp0
        Input #4, tmp1
        Input #5, tmp2
        Input #6, tmp3
        Input #7, tmp4
        MP3Data_C(i) = CInt((CSng(tmp0) + CSng(tmp1) + CSng(tmp2) + CSng(tmp3) + CSng(tmp4)) * 0.2)
         
   Next i

Close #3
Close #4
Close #5
Close #6
Close #7

 
 '====================================================================

Open App.Path & "\MP3_C1.txt" For Input As #3
Open App.Path & "\MP31_C1.txt" For Input As #4
Open App.Path & "\MP32_C1.txt" For Input As #5
Open App.Path & "\MP33_C1.txt" For Input As #6
Open App.Path & "\MP34_C1.txt" For Input As #7




   For i = 0 To 99
    
        Input #3, tmp0
        Input #4, tmp1
        Input #5, tmp2
        Input #6, tmp3
        Input #7, tmp4
        MP3Data_C1(i) = CInt((CSng(tmp0) + CSng(tmp1) + CSng(tmp2) + CSng(tmp3) + CSng(tmp4)) * 0.2)
         
   Next i

Close #3
Close #4
Close #5
Close #6
Close #7



  
 '====================================================================

Open App.Path & "\MP3_C2.txt" For Input As #3
Open App.Path & "\MP31_C2.txt" For Input As #4
Open App.Path & "\MP32_C2.txt" For Input As #5
Open App.Path & "\MP33_C2.txt" For Input As #6
Open App.Path & "\MP34_C2.txt" For Input As #7




   For i = 0 To 99
    
        Input #3, tmp0
        Input #4, tmp1
        Input #5, tmp2
        Input #6, tmp3
        Input #7, tmp4
        MP3Data_C2(i) = CInt((CSng(tmp0) + CSng(tmp1) + CSng(tmp2) + CSng(tmp3) + CSng(tmp4)) * 0.2)
         
   Next i

Close #3
Close #4
Close #5
Close #6
Close #7


'=================================================================
'========= AU3130BLF20 1st mode
Open App.Path & "\MP3_BL.txt" For Input As #3
 




   For i = 0 To 99
    
        Input #3, tmp0
         MP3Data_BL(i) = CInt(CSng(tmp0))
         
   Next i

Close #3
 
'=================================================================
'========= AU3130BLF20 1st mode
Open App.Path & "\MP3_BL1.txt" For Input As #3
 




   For i = 0 To 99
    
        Input #3, tmp0
         MP3Data_BL1(i) = CInt(CSng(tmp0))
         
   Next i

Close #3


'========= AU3130BLF20 1st mode
Open App.Path & "\MP3_CL.txt" For Input As #3

   For i = 0 To 99
    
        Input #3, tmp0
         MP3Data_CL(i) = CInt(CSng(tmp0))
         
   Next i

Close #3

'==================================================
  Open App.Path & "\MP3_CW1.txt" For Input As #3
   For i = 0 To 99
    
        Input #3, tmp0

          MP3Data_CW1(i) = CInt(CSng(tmp0))
         
   Next i
Close #3

'============================================

   Open App.Path & "\MP3_3150J.txt" For Input As #3
    For i = 0 To 99
    
        Input #3, tmp0

          MP3Data_3150J(i) = CInt(CSng(tmp0))
         
    Next i
 Close #3
 
 '============================================

   Open App.Path & "\MP3_3150J1.txt" For Input As #3
    For i = 0 To 99
    
        Input #3, tmp0

          MP3Data_3150J1(i) = CInt(CSng(tmp0))
         
    Next i
 Close #3


'============================================

   Open App.Path & "\MP3_3152A1.txt" For Input As #3
    For i = 0 To 99
    
        Input #3, tmp0

          MP3Data_3152A1(i) = CInt(CSng(tmp0))
         
    Next i
 Close #3

'============================================

   Open App.Path & "\MP3_3152A2.txt" For Input As #3
    For i = 0 To 99
    
        Input #3, tmp0

          MP3Data_3152A2(i) = CInt(CSng(tmp0))
         
    Next i
 Close #3
 
 '============================================
   Open App.Path & "\MP3_3152AL23.txt" For Input As #3
    Open App.Path & "\MP3_3152AL231.txt" For Input As #4
      Open App.Path & "\MP3_3152AL232.txt" For Input As #5
    For i = 0 To 99
    
        Input #3, tmp0
         Input #4, tmp1
          Input #5, tmp2
          MP3Data_3152A3(i) = CInt((CSng(tmp0) + CSng(tmp1) + CSng(tmp2) + CSng(tmp2)) * 0.25)
         
    Next i
    Close #5
 Close #4
  Close #3
'==============================================
' for AU3150ALF22,AU3150ALF22
'============================================

   Open App.Path & "\MP3_3150ALF221.txt" For Input As #3
    For i = 0 To 99
    
        Input #3, tmp0

          MP3Data_3150A221(i) = CInt(CSng(tmp0))
         
    Next i
 Close #3

'============================================

   Open App.Path & "\MP3_3150ALF221.txt" For Input As #3
    For i = 0 To 99
    
        Input #3, tmp0

          MP3Data_3150A222(i) = CInt(CSng(tmp0))
         
    Next i
 Close #3




' for AU3150ALF22,AU3150ALF22
'============================================

   Open App.Path & "\MP3_3150ALF221.txt" For Input As #3
    For i = 0 To 99
    
        Input #3, tmp0

          MP3Data_3150A221(i) = CInt(CSng(tmp0))
         
    Next i
 Close #3

'============================================

   Open App.Path & "\MP3_3150ALF221.txt" For Input As #3
    For i = 0 To 99
    
        Input #3, tmp0

          MP3Data_3150A222(i) = CInt(CSng(tmp0))
         
    Next i
 Close #3
'==============================================

' for AU3150AKL ,
'============================================

   Open App.Path & "\MP3_3150KL1.txt" For Input As #3
    For i = 0 To 99
    
        Input #3, tmp0

          MP3Data_3150KL1(i) = CInt(CSng(tmp0))
         
    Next i
 Close #3

'============================================

   Open App.Path & "\MP3_3150KL2.txt" For Input As #3
    For i = 0 To 99
    
        Input #3, tmp0

          MP3Data_3150KL2(i) = CInt(CSng(tmp0))
         
    Next i
 Close #3
'==============================================

 'for AU3150BKL ,
'============================================

   Open App.Path & "\MP3_3150KL21.txt" For Input As #3
    For i = 0 To 99
    
        Input #3, tmp0

          MP3Data_3150KL21(i) = CInt(CSng(tmp0))
         
    Next i
 Close #3

'============================================

   Open App.Path & "\MP3_3150KL22.txt" For Input As #3
    For i = 0 To 99
    
        Input #3, tmp0

          MP3Data_3150KL22(i) = CInt(CSng(tmp0))
         
    Next i
 Close #3
'==============================================

  Open App.Path & "\MP3_3150KL23.txt" For Input As #3
    For i = 0 To 99
    
        Input #3, tmp0

          MP3Data_3150KL23(i) = CInt(CSng(tmp0))
         
    Next i
 Close #3
'==============================================
'==============================================

  Open App.Path & "\MP3_3150KL23WMA11.txt" For Input As #3
    For i = 0 To 99
    
        Input #3, tmp0

          MP3WMA_3150KL23(i) = CInt(CSng(tmp0))
         
    Next i
 Close #3
 
 
'==============================================


'==============================================

  Open App.Path & "\MP3_3150KL23WMA12.txt" For Input As #3
    For i = 0 To 99
    
        Input #3, tmp0

          MP3WMA_3150KL231(i) = CInt(CSng(tmp0))
         
    Next i
 Close #3
 
 
'==============================================

'==============================================

  Open App.Path & "\MP3_3150KL23WMA13.txt" For Input As #3
    For i = 0 To 99
    
        Input #3, tmp0

          MP3WMA_3150KL232(i) = CInt(CSng(tmp0))
         
    Next i
 Close #3
 
 
'==============================================


 Open App.Path & "\MP3_AU62541.txt" For Input As #3
    For i = 0 To 99
    
        Input #3, tmp0

          MP3Data_AU6254(i) = CInt(CSng(tmp0))
         
    Next i
 Close #3

 

'==============================================


 Open App.Path & "\MP3_AU62542.txt" For Input As #3
    For i = 0 To 99
    
        Input #3, tmp0

          MP3Data_AU62541(i) = CInt(CSng(tmp0))
         
    Next i
 Close #3
'==============================================

 

Open App.Path & "\SCSI0.txt" For Input As #3
Open App.Path & "\SCSI1.txt" For Input As #4
Open App.Path & "\SCSI2.txt" For Input As #5
Open App.Path & "\SCSI3.txt" For Input As #6




   For i = 0 To 35
    
        Input #3, InquiryString(0, i)
        Input #4, InquiryString(1, i)
        Input #5, InquiryString(2, i)
        Input #6, InquiryString(3, i)
      
   Next i

Close #3
Close #4
Close #5
Close #6
 

'============ MMC control initail ==================
    mmcAudio.Notify = False
    mmcAudio.Wait = True
    mmcAudio.Shareable = False
    mmcAudio.Command = "Close"
GetVol



Dim lngResult As Long

    bReadyToClose = False
    bCurSlotNum = 0
    strAppPath = App.Path
    
    If (Right(strAppPath, 1) <> "\") Then
        strAppPath = strAppPath & "\"
    End If

    ' Connect to the smartcard resource manager
    lngResult = SCardEstablishContext(SCARD_SCOPE_SYSTEM, lngNull, lngNull, lngContext)
    
     If (lngResult <> 0) Then
     '    MsgBox ("SCardEstablishContext Failed")
     '     GoTo Error_Exit
      End If

    
    IsReaderLost = True
    lngResult = ConnectAlcorReader
   
    
    SaveOSCounter = 1
  


GPIBCard_Exist = False
GPIBCard_Exist = CheckGPIB()
ResetHubString = App.Path & "\devcon restart @USB\ROOT_HUB*"

SkipCtrlCount = "0"

If GPIBCard_Exist = False Then
    GPIBCARD_Label.Caption = "  No GPIB "
    GPIBCARD_Label.ForeColor = &HFF&
    SkipCtrlCount.Enabled = False
    
    If Dir("D:\NoGPIB.PC") = "" Then
        Open "D:\NoGPIB.PC" For Output As #55
        Call MsecDelay(0.02)
        Close #55
    End If
Else
    GPIBCARD_Label.Caption = " GPIB Exist"
    GPIBCARD_Label.ForeColor = &HFF0000
    If Dir("D:\NoGPIB.PC") = "NoGPIB.PC" Then
        Kill ("D:\NoGPIB.PC")
    End If
End If

If Dir("D:\OSFail_Log", vbDirectory) = "" Then
    MkDir ("D:\OSFail_Log\")
End If

RunReleaseMemCount = 0

Exit Sub

Error_Exit:
End


End Sub

Private Sub Form_Activate()
    Call Command1_Click
End Sub
Private Sub Command1_Click()
 On Error Resume Next
Const AllenTest = 0
Dim i As Integer
Dim AU6981FailCounter As Integer
Dim TestCounter As Integer
Dim OldTimer
Dim receivestart As String
Dim goodchip As Long
Dim failchip As Long
Dim failchipBin2 As Long 'arch add 940411
Dim failchipBin3 As Long 'arch add 940411
Dim failchipBin4 As Long 'arch add 94041133
Dim failchipBin5 As Long 'arch add 940411
Dim Endmsg As String
Dim LastCardFail  As Byte
 
Dim DirName As String
Dim LogFile As String
Dim Counter As Integer
Dim f1 As String
Dim f2 As String
Dim v1 As Byte
Dim v2 As Byte
Dim v3 As Byte
Dim OldLBa As Long


Dim AllChip As String

 
Dim OldTime
Dim FailPosition As Integer

'Const UNKNOW_DEVICE = 0

Dim TmpLBA  As Long
Dim TesterReadyCounter As Integer

Dim OldTestResult As String
Dim GPIBStatus As String

Const Ver = " v1.35"
'//////////////// need modify

'ChDir ("D:\")

AllChip = "AU6330"
AllChip = AllChip & "AU6331"
AllChip = AllChip & "AU6333"
AllChip = AllChip & "AU6363"
AllChip = AllChip & "AU6366"
AllChip = AllChip & "AU6367"
AllChip = AllChip & "AU6368"
AllChip = AllChip & "AU6369_1"
AllChip = AllChip & "AU6369_2"
AllChip = AllChip & "AU6369S2"
AllChip = AllChip & "AU6369S3"
AllChip = AllChip & "AU6375"
AllChip = AllChip & "AU6384"
AllChip = AllChip & "AU6385"
AllChip = AllChip & "AU6386"
AllChip = AllChip & "AU6386_D"
AllChip = AllChip & "AU6390"
AllChip = AllChip & "AU9368"
AllChip = AllChip & "AU9369S3"
AllChip = AllChip & "AU9368_1"
AllChip = AllChip & "AU9368_S"
AllChip = AllChip & "AU9510"
AllChip = AllChip & "AU9520"
AllChip = AllChip & "AU9720"


'\\\\\\\\\\\\\\\\\\\\\\\\\\\
' initial software labeling
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Label17.Caption = "UNKNOW DEVICE"
Label18.Caption = "NO DEFINE"
Label19.Caption = "NO DEFINE"
Label20.Caption = "NO DEFINE"

Label21.Caption = "SD WRITE FAIL"
Label22.Caption = "SD READ FAIL"
Label23.Caption = "CF WRITE FAIL"
Label24.Caption = "CF READ FAIL"

Label25.Caption = "XD WRITE FAIL"
Label26.Caption = "XD READ FAIL"
Label27.Caption = "NO DEFINE"
Label28.Caption = "NO DEFINE"

Label29.Caption = "MS WRITE FAIL"
Label30.Caption = "MS READ FAIL"
Label31.Caption = "NO DEFINE"
Label32.Caption = "NO DEFINE"

Label12 = "Testing...."

Label12.BackColor = &H8080FF
Do
    
     OldChipName = ""
     ChipName = ""
     DoEvents
     UsbSpeedTestResult = 0
     FlashCapacityError = 0
     GPOFail = 0
     bNeedsReStart = False
     
     rv0 = 0
     rv1 = 0
     rv2 = 0
     rv3 = 0
     rv4 = 0
     rv5 = 0
     rv6 = 0
     rv7 = 0
     
    
    
         
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '
    ' Wait for begin Test command
    '
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        
GPIB_FLAG:

        ChipName = ""
        
        If AtheistDebug Then
            ChipName = AtheistDebugName
        End If
        
         'Do
            '   MSComm1.Output = "Ready"
            '   TesterReadyCounter = TesterReadyCounter + 1
               'Print TesterReadyCounter
             '  DoEvents
             '  Call MsecDelay(0.1) ' comm time
             '  ChipName = MSComm1.Input
             ' fnScsi2usb2K_KillEXE ' clear removable device message box
              ' ChipName = "AU9368_1"
        ' Loop Until ChipName <> ""
         
         
        
         MSComm1.InBufferCount = 0
         MSComm1.OutBufferCount = 0
        Print "wait host comannd"
             
        Do
           
               MSComm1.Output = "Ready"
               Call MsecDelay(0.01)
               
               DoEvents
     
     
               buf = MSComm1.Input
               ChipName = ChipName & buf
              fnScsi2usb2K_KillEXE ' clear removable device message box
             '     ChipName = "AU6395BLF20"
         Loop Until InStr(1, ChipName, "AU") <> 0 And NameLen = 1
         
        If InStr(1, ChipName, "AUGPIB") <> 0 Then
           If SkipCtrlCount.Enabled = True Then
               GPIBStatus = "GPIBReadyGPIBReady"
           ElseIf SkipCtrlCount.Enabled = False Then
               GPIBStatus = "GPIBUNReadyGPIBUNReady"
           End If
               
           Do
               MSComm1.Output = GPIBStatus
               Call MsecDelay(0.02)
               DoEvents
               buf = MSComm1.Input
               ChipName = ChipName & buf
           Loop Until (InStr(1, ChipName, "AUGPIBACK") <> 0)
           GoTo GPIB_FLAG
        End If
        
          
          ReaderExist = 0
          If InStr(ChipName, "~") <> 0 Then
            Alarm_NonUPT2.Show
            Exit Sub
          End If
          
          ChipName = Right(ChipName, Len(ChipName) - InStr(1, ChipName, "AU") + 1)
         
          Cls
          
           Label1.Caption = ChipName & " Tester" & Ver
           Print "Get Host Command"
        '///////////////////////////////
        '
        '  Begin Test
        '
        '///////////////////////////////
         
         '=====  Initial all HMI interface
         
         TestResult = "FAIL"
         Label9 = ""
         Label3.BackColor = RGB(255, 255, 255)
         Label4.BackColor = RGB(255, 255, 255)
         Label5.BackColor = RGB(255, 255, 255)
         Label6.BackColor = RGB(255, 255, 255)
         Label7.BackColor = RGB(255, 255, 255)
         Label8.BackColor = RGB(255, 255, 255)
         
         
        '======= Test loop
         
         
        If SkipCtrlCount.Enabled = True Then
            If SkipCtrlCount = "0" Then
                If Dir("D:\NoGPIB.PC") = "NoGPIB.PC" Then
                    Kill ("D:\NoGPIB.PC")
                End If
            Else
                SkipCtrlCount = CStr(CInt(SkipCtrlCount) - 1)
                If Dir("D:\NoGPIB.PC") = "" Then
                    Open "D:\NoGPIB.PC" For Output As #55
                    Call MsecDelay(0.02)
                    Close #55
                End If
            End If
        End If
         
        Print "Begin Test"
                
        OldTime = Timer
        
        Select Case ChipName
                Case "AU3510ELF20", "AU3510ELF21", "AU3510ELF21"
                    Call AU3510ELF20TestSub
                    
                Case "AU3522DLF20", "AU3522DLF21"
                    Call AU3522DLF20TestSub
                    
'                Case "AU3510ENG20"
'                    Call AU3510ENG20TestSub
                    
                 Case "AU9562GFF20"
                    Call AU9562GFF20TestSub
            
                 Case "AU1111AAA10"
                 
                    Call AU1111AAA10TestSub
                    
                 Case "AU6485AFP10", "AU6485BFP10", "AU6485CFP10", "AU6485DFP10", "AU6485HFP10", "AU6485IFP10", "AU6485JFP10"
                    
                    Call OpenShortTest_AssignOSFileName("AU6485AFP10")
                    
                 Case "AU9562BSP10", "AU9562AFP10", "AU8451DBP10", "AU8451EBP10", "AU6601BFP10", "AU6601CFP10", "AU6621CFP10", "AU6621CFP11", "AU6621GFP10", "AU6621DFP10", "AU6621BFP10", "AU2101AFP11", "AU2101BFP11", "AU2101DFP11", "AU2101EFP11", "AU2101HFP11", "AU2101ASP10", "AU2101BSP10"
                    Call OpenShortTest_SkipZero
        
                 Case "AU6433BSP10", "AU6433BLP10", "AU6433DLP10", "AU6433DLP1A", "AU6473CLP10", "AU6473BLP10", "AU6992DLP10", "AU6376ALP10", "AU6433JSP10", "AU6433EFP10", "AU6433LFP10", "AU6429FLP10", "AU6425DLP10", "AU6427ELP10", "AU6427GLP10"
                     
                    Call OpenShortTest
                    
                 Case "AU6435ELP10", "AU6435ELP1A", "AU6435GLP1A", "AU6435BFP10", "AU6435CFP10", "AU6913DLP10", "AU6915MLP10", "AU6916DLP10", "AU6257ELP10", "AU6915DLP10", "AU6259BLP10", "AU6259BFP10", "AU6259CLP10", "AU6917OLP10", "AU6257BFP10"
                    
                    Call OpenShortTest
                    
                 Case "AU6438IFP10", "AU6479BLP10", "AU6479NLP10", "AU6479HLP10", "AU6479FLP10", "AU6479ALP10", "AU6479ILP10", "AU6479JLP10", "AU6479KLP10", "AU6479OLP10", "AU6479TLP10", "AU6479BFP10", "AU8451DBP10", "AU6479CFP10", "AU6479DFP10", "AU6479VLP10", "AU6479ULP10"
                 
                    Call OpenShortTest
                    
                Case "AU6438BSP10"
                    
                    Call OpenShortTest_SkipZero_NoGPIB
                    
                Case "AU6259ILP10"
                
                    Call OpenShortTest
                    
                Case "AU6366CLP10"
                
                    Call OpenShortTest
                    
                 Case "AU6485LFP10", "AU6350ELP10"
                
                    Call OpenShortTest
                    
                 Case "AU6910DLP10", "AU6919DLP10", "AU6991DLP10", "AU6922DLP10", "AU6922OLP10"
                 
                    Call OpenShortTest
                
                Case "AU6921DLP10"
                    Call OpenShortTest
                    
                 Case "AU6919DLP20"
                 
                    Call OpenShortTest_Pin2ShortBin3
                    
                 Case "AU6990DLP11", "AU6990DLP1A"  '11: GSMC chip ; 1A: CSMC chip
                    
                    Call OpenShortTest
                 
                 Case "AU6992DLP11", "AU6992DLP1A"  '11: GSMC chip ; 1A: CSMC chip
                    
                    Call OpenShortTest
                 
                 Case "AU6990DLS10"
                    
                    Call AU6990HW_SortTest
                    
                Case "AU6996DLP10"
                    Call OpenShortTest
                    
                Case "AU6922DLS10"
                    
                    Call AU6922HW_SortTest
                
                Case "AU6601CFF20", "AU6601CFF00"
                
                    Call AU6601FTTestSub
                
                
                 Case "AU6350KLF21", "AU6350KLF22", "AU6350KLF23", "AU6350ALF20", "AU6350BFF20", "AU6350GLF20", "AU6350OLF21", "AU6350ALF21", "AU6350ALF22", "AU6350BFF21", "AU6350GLF21", "AU6350BLF21", "AU6350CFF21", "AU6350CFF22"
                       
                        Call AU6350Test
                        
                  Case "AU6710ASF20"
                  
                        Call AU6710ASTest
        
                  Case "AU6395CLS11", "AU6395CLS10", "AU6395CLF20", "AU6395CLF21"
                  
                       If ChipName = "AU6395CLF20" Then
                       
                        Call AU6395CLTestSub
                       End If
                       
                      If ChipName = "AU6395CLF21" Then
                        Call AU6395CLF21TestSub
                       End If
                        
                      If ChipName = "AU6395CLS10" Then
                        Call AU6395CLS10SortingSub
                       End If
                        
                       If ChipName = "AU6395CLS11" Then
                        Call AU6395CLS11SortingSub
                       End If
                        
                    Case "AU6395BLF20"
                  
                        Call AU6395BLTestSub
                        
                    Case "AU6366C"
                        
                        Call AU6366C_IQCSub
                    
                  Case "AU6376ILF21"
                  
                       Call MultiSlotTestAU6376
                       
                  
                   Case "AU6254XLS40", "AU6254XLS41"
             
                       If ChipName = "AU6254XLS40" Then
                       
                       
                        Call AU6254CMedia
                        
                       End If
                
                        If ChipName = "AU6254XLS41" Then
                       
                       
                        Call AU6254CMedia1
                        
                       End If
                 
                      Case "AU6982HLS10"
                      
                    
                        If PCI7248InitFinish = 0 Then
                          PCI7248Exist
                        End If
                        
                        
                         '=========================================
                        '    POWER on
                        '=========================================
                         CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                         
                         If CardResult <> 0 Then
                            MsgBox "Power off fail"
                            End
                         End If
                         
                         Call MsecDelay(0.05)
                         
                         CardResult = DO_WritePort(card, Channel_P1A, &H0)  'Power Enable
                         Call MsecDelay(1.8)    'power on time
                         
                          If CardResult <> 0 Then
                            MsgBox "Power on fail"
                            End
                         End If
                         
                      
                 
                          rv0 = CBWTest_New_AU6390MB(0, 1, "vid_058f", 0)
                          
                       
                       
                         LBA = 0
                         For i = 1 To 31
                             rv1 = 0
                             LBA = LBA + 254976
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector(0, 1, 1)  ' write
                             If rv1 <> 1 Then
                             GoTo AU6982HLS10Result
                             End If
                         Next
                         
                         LBA = 0
                         For i = 1 To 31
                            rv2 = 0
                            LBA = LBA + 254976
                            ClosePipe
                            rv2 = CBWTest_New_128_Sector(0, 1, 0)  ' read
                            If rv2 <> 1 Then
                            GoTo AU6982HLS10Result
                            End If
                         Next
                        
                        
                        
                        
AU6982HLS10Result:
        
                       
                       If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv0 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                ElseIf rv1 = WRITE_FAIL Then
                    CFWriteFail = CFWriteFail + 1
                    TestResult = "CF_WF"
                ElseIf rv1 = READ_FAIL Then
                    CFReadFail = CFReadFail + 1
                    TestResult = "CF_RF"
                ElseIf rv2 = WRITE_FAIL Then
                    XDWriteFail = XDWriteFail + 1
                    TestResult = "XD_WF"
                ElseIf rv2 = READ_FAIL Then
                    XDReadFail = XDReadFail + 1
                    TestResult = "XD_RF"
                 ElseIf rv3 = WRITE_FAIL Then
                    MSWriteFail = MSWriteFail + 1
                    TestResult = "MS_WF"
                ElseIf rv3 = READ_FAIL Then
                    MSReadFail = MSReadFail + 1
                    TestResult = "MS_RF"
                ElseIf rv0 * rv1 * rv2 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
                 
  
        
                  Case "AU6981HLF27", "AU6981DLF20"
                   
                   
                      
              If Left(ChipName, 10) = "AU6981HLF2" Then
                  ChipNo = 1
              End If
              
                If Left(ChipName, 10) = "AU6981DLF2" Then
                  ChipNo = 0
              End If
                   
                   
                 If LastCardFail = 1 Or FixCardMode.Value = 1 Then
                 Print "Fix Card"
                 Call FixCardSub
                
                 End If
                   
               
                   
           
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                End If
                
                
                 '=========================================
                '    POWER on
                '=========================================
                 CardResult = DO_WritePort(card, Channel_P1A, &HFF) 'ENA_B=0, ENA_A=0, SEL=1
                 
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                  If LastCardFail = 1 Or FixCardMode.Value = 1 Then
                   Call MsecDelay(2)
                  LastCardFail = 0
                  Else
                  Call MsecDelay(0.5)
                  End If
                  
                 CardResult = DO_WritePort(card, Channel_P1A, &HFE)  'ENA_B=0, ENA_A=1, SEL=0
                 
                 Call MsecDelay(0.5)
                 result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
                 CardResult = DO_WritePort(card, Channel_P1A, &HF2)  'ENA_B=0, ENA_A=1, SEL=0
                 Call MsecDelay(1.8)    'power on time
                result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                  If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                 End If
                 
            LightOn = 0
                 CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                   Call MsecDelay(0.5)
                 
               LBA = LBA + 1   ' for dual channel
               
               Print LBA
               
                ClosePipe
                ReaderExist = 0
                rv0 = CBWTest_New_AU6390MB(0, 1, "vid_058f", 0)
                 Call LabelMenu(0, rv0, 1)
                ClosePipe
                rv1 = CBWTest_New_AU6390MB(0, rv0, "vid_058f", 1)
                Call LabelMenu(1, rv1, rv0)
                 ClosePipe
                  OldLBa = LBA
                  
                  LBA = LBA + 4079614 'org= 3932160=7C7FFD*0.5
                  LBA = LBA + 65535
                 ClosePipe
                 rv2 = CBWTest_New_AU6390MB(0, rv1, "vid_058f", 2)
                 ClosePipe
                Call LabelMenu(2, rv2, rv1)
                
                '============ AU6982IL  the second 2 MBG Flash begin ==========
                If ChipName = "AU6982IL" And rv2 = 1 Then
                
                    LBA = LBA + 4079614 'org= 3932160=7C7FFD*0.5
                    LBA = LBA + 65535
                   ClosePipe
                   rv2 = CBWTest_New_AU6390MB(0, rv1, "vid_058f", 2)
                   ClosePipe
                   Call LabelMenu(2, rv2, rv1)
                End If
                
                 If ChipName = "AU6982IL" And rv2 = 1 Then
                
                    LBA = LBA + 4079614 'org= 3932160=7C7FFD*0.5
                    LBA = LBA + 65535
                    ClosePipe
                   rv2 = CBWTest_New_AU6390MB(0, rv1, "vid_058f", 2)
                   ClosePipe
                   Call LabelMenu(2, rv2, rv1)
                End If
                 '============ AU6982IL  the second 2 MBG Flash begin ==========
                
                 LBA = OldLBa
                
                ClosePipe
                 Call MsecDelay(0.5)
                
                   If CardResult <> 0 Then
                    MsgBox "Read card detect light ON fail"
                    End
                   End If
                   
                   
                 If ChipName = "AU6980MB" Or "AU6981HLF25" Then
                 Print "Light="; LightOn
                       If rv2 = 1 Then
                          If LightOn = 207 Then
                             rv3 = 1
                           Else
                             GPOFail = 2
                             rv3 = 2
                          End If
                       Else
                            rv3 = 4
        
                       End If
                End If
                 
                 
                   If ChipName = "AU6982IL" Then
                       If rv2 = 1 Then
                          If LightOn = 254 Then
                             rv3 = 1
                           Else
                             GPOFail = 2
                             rv3 = 2
                          End If
                       Else
                            rv3 = 4
        
                       End If
                End If
                 
                 
                  Call LabelMenu(32, rv3, rv2)
                 
                    
                   CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                 
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                  
                 Print "rv0=;"; rv0
                 Print "rv1=;"; rv1
                 Print "rv2=;"; rv2
                 
                  
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv0 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                ElseIf rv1 = WRITE_FAIL Then
                    CFWriteFail = CFWriteFail + 1
                    TestResult = "CF_WF"
                ElseIf rv1 = READ_FAIL Then
                    CFReadFail = CFReadFail + 1
                    TestResult = "CF_RF"
                ElseIf rv2 = WRITE_FAIL Then
                    XDWriteFail = XDWriteFail + 1
                    TestResult = "XD_WF"
                ElseIf rv2 = READ_FAIL Then
                    XDReadFail = XDReadFail + 1
                    TestResult = "XD_RF"
                 ElseIf rv3 = WRITE_FAIL Then
                    MSWriteFail = MSWriteFail + 1
                    TestResult = "MS_WF"
                ElseIf rv3 = READ_FAIL Then
                    MSReadFail = MSReadFail + 1
                    TestResult = "MS_RF"
                ElseIf rv0 * rv1 * rv2 * rv3 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
               
                If TestResult <> "PASS" And rv0 <> UNKNOW Then
                
                If rv0 * rv1 * rv2 <> 1 Then
                
                LastCardFail = 1
                End If
                End If
                 Case "AU6981HLF24"
          
  
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                End If
                
                
                 '=========================================
                '    POWER on
                '=========================================
                 CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                 
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 
                 Call MsecDelay(0.05)
                 CardResult = DO_WritePort(card, Channel_P1A, &H0)  'Power Enable
                 Call MsecDelay(1.8)    'power on time
                 
                  If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                 End If
                 
              ' ======== Recovery ===== function
              
               rv0 = AU6981_Recovery_Initial(0, 1, "vid", 0)
                 Call LabelMenu(0, rv0, 1)
               If rv0 <> 1 Then
                     'MsgBox "card detect fail"
                     GoTo TestResultLabel_AU6981
               End If
               ClosePipe
               OpenPipe
               rv1 = AU6981_EraseBlock0Test
               ClosePipe   'D0 00 F0 60 F1 03 00 00 00 F0 D0 F0 70 F3 01 00 >> 60 start 至70 end,位址03 00 00 00,F3讀 1 byte
                 Call LabelMenuSu(1, rv1, rv0)
               If rv1 <> 1 Then
                    'MsgBox "EraseBlock0 fail"
                    GoTo TestResultLabel_AU6981
               End If
               ClosePipe
                OpenPipe
               rv2 = AU6981_WritePhsyicalTest1 'D0 00 F0 80 F1 05 00 00 00 00 00 F5 01 03 F7 00 >> write命令 80 start 位址05 00 00 00 00 00,write sector 01 03,single 0
               ClosePipe
                
               Call LabelMenuSu(2, rv2, rv1)
               If rv2 <> 1 Then
                    'MsgBox "card write fail"
                    GoTo TestResultLabel_AU6981
               End If
               ClosePipe
                OpenPipe
               rv3 = AU6981_WritePhsyicalTest2 'D0 00 F0 10 F0 70 F3 01 05 00 00 00 00 00 00 00 >> 10 start 至70 end,F3命令讀1 byte,位址05 00 00 00 00 00
               
              ClosePipe
                 Call LabelMenuSu(3, rv3, rv2)
               'If rv3 <> 1 Then
                    'MsgBox "card writr fail"
                   ' GoTo TestResultLabel_AU6981
              ' End If
               ClosePipe
                OpenPipe
               rv4 = AU6981_ReadPhsyicalTest   'D0 00 F0 00 F1 05 00 00 00 00 00 F0 30 F4 01 00 >> start 00,位址05 00 00 00 00 00,read sector 01
               ClosePipe
               Call LabelMenuSu(4, rv4, rv3)
               If rv4 <> 1 Then
                    'MsgBox "card read fail"
                    GoTo TestResultLabel_AU6981
               End If
              '====================================
                   
                   LightOn = 0
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                  
                   If CardResult <> 0 Then
                    MsgBox "Read card detect light ON fail"
                    End
                   End If
                   
                   
                
                    If rv4 = 1 Then
                          If LightOn = 223 Then
                             rv5 = 1
                           Else
                             GPOFail = 2
                             rv5 = 2
                          End If
                    
        
                    End If
                 
                   'Call LabelMenu(32, rv3, rv2)
                   Call LabelMenuSu(4, rv5, rv4)
                     
                   CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                 
                   If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                   End If
                  
 
TestResultLabel_AU6981: Print "rv0="; rv0
                        Print "rv1="; rv1
                        Print "rv2="; rv2
                        Print "rv3="; rv3
                        Print "rv4="; rv4
                         Print "rv5="; rv5

                        If rv0 = UNKNOW Then
                           UnknowDeviceFail = UnknowDeviceFail + 1
                           TestResult = "UNKNOW"
                          
                        ElseIf rv0 = WRITE_FAIL Then
                           SDWriteFail = SDWriteFail + 1
                           TestResult = "SD_WF"
                           
                        ElseIf rv0 = READ_FAIL Then
                           SDReadFail = SDReadFail + 1
                           TestResult = "SD_RF"
                        ElseIf rv1 = WRITE_FAIL Then
                           CFWriteFail = CFWriteFail + 1
                           TestResult = "CF_WF"
                        ElseIf rv1 = READ_FAIL Then
                           CFReadFail = CFReadFail + 1
                           TestResult = "CF_RF"
                        ElseIf rv2 = WRITE_FAIL Then
                           XDWriteFail = XDWriteFail + 1
                           TestResult = "XD_WF"
                        ElseIf rv2 = READ_FAIL Then
                           XDReadFail = XDReadFail + 1
                           TestResult = "XD_RF"
                        ElseIf rv3 = WRITE_FAIL Or rv4 = WRITE_FAIL Or rv5 = WRITE_FAIL Then
                           MSWriteFail = MSWriteFail + 1
                           TestResult = "MS_WF"
                        ElseIf rv3 = READ_FAIL Or rv4 = READ_FAIL Or rv5 = READ_FAIL Then
                           MSReadFail = MSReadFail + 1
                           TestResult = "MS_RF"
                        ElseIf rv0 * rv1 * rv2 * rv3 * rv4 * rv5 = 1 Then
                           TestResult = "PASS"
                       
                        Else
                           TestResult = "Bin2"
                        End If
        
        
              
           
     
                
        
           Case "AU7630"
                  
                  
                If ChipName = "AU7630" Then
                   If PCI7248InitFinish = 0 Then
                      PCI7248Exist
                    End If
                    
                    CardResult = DO_WritePort(card, Channel_P1A, &H0)
                    Call MsecDelay(1#)
                     
                End If
                LBA = LBA + 1
                
                
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
                
                
                Print LBA
                
                ClosePipe
                rv0 = CBWTest_New(0, 1, "vid_058f")
                'Print "a1"
                Call LabelMenu(0, rv0, 1)
                ClosePipe
               
                
                
                
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv0 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                ElseIf rv0 = PASS Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                End If
                
                Print "Test Result"; TestResult
                       
           
           
        
            
           Case "AU9368_B", "AU9368ALF20"
                
                If ChipName = "AU9368_B" Then
                 Call AU9368BTest
                End If
                
                
                If ChipName = "AU9368ALF20" Then
                 Call AU9368ALFTest
                End If
        
        
          Case "AU9331", "AU6254BLS50"
          
             If ChipName = "AU6254BLS50" Then
                       Call AU6254BLS50SortingSub
                       If rv0 = UNKNOW Then
                           UnknowDeviceFail = UnknowDeviceFail + 1
                           TestResult = "UNKNOW"
                        Else
                        
                           TestResult = "PASS"
                        End If
               ElseIf ChipName = "AU9331" Then
 
                  Call AU9331Test
              
              End If
                 
                
           
           Case "AU6254", "AU6254ALF20", "AU6254BLF20", "AU6254ALF22", "AU6254BLF22", "AU6254DLF22", "AU6254AFF20", "AU6254XLT20", "AU6254BLF23", "AU6254BLF23", "AU6256BLF20", "AU6254ALF23", "AU6254ASF22", "AU6256CFF22"
        
          If ChipName = "AU6254" Or ChipName = "AU6254ALF20" Or ChipName = "AU6254BLF20" Then
          
           Call AU6254Test(rv0, rv1, rv2, rv3, rv4)
           
        
          Else
          
            
             CardResult = DO_WritePort(card, Channel_P1A, &H1F)
              Call MsecDelay(0.1)
              
            If ChipName = "AU6254BLF23" Or ChipName = "AU6254ALF23" Then
              Call AU6254TestAlcorV2(rv0, rv1, rv2, rv3, rv4)
            ElseIf ChipName = "AU6256BLF20" Then
               Call AU6254TestAlcorAU6256BLF20(rv0, rv1, rv2, rv3, rv4)
            ElseIf ChipName = "AU6254ASF22" Or ChipName = "AU6256CFF22" Then
               Call AU6254TestAlcorAU6254AS(rv0, rv1, rv2, rv3, rv4)
            Else
              Call AU6254TestAlcor(rv0, rv1, rv2, rv3, rv4)
            End If
            
            
            CardResult = DO_WritePort(card, Channel_P1A, &H1F)
          End If
           
           
                Print " test hub disconnect -------"
               ' CardResult = DO_WritePort(card, Channel_P1A, &HFF)
              
                
                Print "rv0="; rv0
                Print "rv1="; rv1
                Print "rv2="; rv2
                Print "rv3="; rv3
                Print "rv4="; rv4
                 
                  
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv0 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                ElseIf rv1 = WRITE_FAIL Then
                    CFWriteFail = CFWriteFail + 1
                    TestResult = "CF_WF"
                 ElseIf rv1 = READ_FAIL Then
                    CFReadFail = CFReadFail + 1
                    TestResult = "CF_RF"
                ElseIf rv3 = WRITE_FAIL Then
                    XDWriteFail = XDWriteFail + 1
                    TestResult = "XD_WF"
                ElseIf rv3 = READ_FAIL Then
                    XDReadFail = XDReadFail + 1
                    TestResult = "XD_RF"
                 ElseIf rv4 = WRITE_FAIL Then
                    MSWriteFail = MSWriteFail + 1
                    TestResult = "MS_WF"
                ElseIf rv4 = READ_FAIL Then
                    MSReadFail = MSReadFail + 1
                    TestResult = "MS_RF"
                ElseIf rv0 * rv1 * rv2 * rv3 * rv4 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
                  
                 
                Call LabelMenu(0, rv0, 1)
                Call LabelMenu(1, rv1, rv0)
                Call LabelMenu(2, rv3, rv1)
                Call LabelMenu(3, rv4, rv3)
                
                
                Label9.Caption = AU6254TestMsg
             
        
         
      '  mmcAudio.Command = "Play"
        
           Case "AU3130", "AU3130B", "AU3130CL", "AU3130BLF20", "AU3130CLF20", "AU3150JLF21", "AU3150ALF21", "AU3150ILF21", "AU3150CLF21", "AU3150LLF21", "AU3152ALF21", "AU3150KLF21", "AU3150KLF22", "AU3150KLF23", "AU3152ALF23"
                 Call AU3130_1_SlotTest
           
            Case "AU3150CLF2C", "AU3150CLF27", "AU3152CLF27", "AU3150KLF24", "AU3152CLF24", "AU3150JLF27", "AU3150ALF27", "AU3150KLF27", "AU3152ALF24", "AU3150ALF24", "AU3150NLF27", "AU3150PLF27", "AU3150QLF27", "AU3152ALF27", "AU3150LLF27"
            
               Call MP3Tester
            Case "AU3150LLF2C", "AU3152CLF2A", "AU3152ALF28", "AU3152CLF28", _
                "AU3150ALF28", "AU3150CLF28", "AU3150JLF28", "AU3150KLF28", "AU3150LLF28", "AU3150MLF28", "AU3150NLF28", "AU3150PLF28", "AU3150QLF28"

                 Call MP3Tester
                 
            Case "AU3152HLF2C", "AU3152ALF2C"
                  Call MP3Tester
                 
            Case "AU3152ALF2D", "AU3150LLF2D", "AU3152HLF2D", "AU3152HLF2E"
                 Call MP3Tester
                 
                 
            Case "AU3150ALF22", "AU3150CLF22"
               Call AU3130_2_SlotTest
           
            Case "AU3150MLF23"
                
                Call AU3150MLTestSub
                
            Case "AU9420BLF30", "AU9420DLF30", "AU9420DLF00"
                
                Call AU9420TestSub
         
            Case "AU9540BSF00"
                Call AU9540BSF00TestSub
                
                bNeedsReStart = True
                RunReleaseMemCount = RunReleaseMemCount + 1
                
                If RunReleaseMemCount >= 10 Then
                    Call ReleaseMem(30)
                    RunReleaseMemCount = 0
                End If
           
            Case "AU9520V5", "AU9520_1", "AU9520V4", "AU9520ALF20"
           
                bNeedsReStart = True
                Call AU9520TestGPIB
                RunReleaseMemCount = RunReleaseMemCount + 1
                
                If RunReleaseMemCount >= 10 Then
                    Call ReleaseMem(30)
                    RunReleaseMemCount = 0
                End If
                
            Case "AU9520FLF21", "AU9520V51", "AU9520ALF21", "AU9525ALF20", "AU9520FLF20", "AU9520GLF20", "AU9520ASF20", "AU9540BSF20"
                
                If ChipName = "AU9520V51" Or ChipName = "AU9520ALF21" Then
                    Call AU9520TestNOGPIB
                End If
                
                If ChipName = "AU9520ASF20" Or _
                   ChipName = "AU9540BSF20" Then
                    Call AU9520ASF20TestSub
                End If
                
                
                If ChipName = "AU9525ALF20" Then
                   Call AU9525ALF20TestNOGPIB
                End If
                
                  If ChipName = "AU9520FLF20" Then
                   Call AU9520FLF20TestSub
                 ElseIf ChipName = "AU9520GLF20" Then
                   Call AU9520GLF20TestSub
                  ElseIf ChipName = "AU9520FLF21" Then
                   Call AU9520FLF21TestSub
                  End If
                
                bNeedsReStart = True
                RunReleaseMemCount = RunReleaseMemCount + 1
                
                If RunReleaseMemCount >= 10 Then
                    Call ReleaseMem(30)
                    RunReleaseMemCount = 0
                End If
            
            Case "AU9562AFE10"
            
                Call AU9562AFE10TestSub
            
            Case "AU9562AFF20", "AU9562BSF20"
                
                Call AU9520ASF20TestSub
                
                bNeedsReStart = True
                RunReleaseMemCount = RunReleaseMemCount + 1
                
                If RunReleaseMemCount >= 10 Then
                    Call ReleaseMem(30)
                    RunReleaseMemCount = 0
                End If
                
            Case "AU9562BSF40"
                
                Call AU9562BSF40TestSub
                
                bNeedsReStart = True
                RunReleaseMemCount = RunReleaseMemCount + 1
                
                If RunReleaseMemCount >= 10 Then
                    Call ReleaseMem(30)
                    RunReleaseMemCount = 0
                End If
                
            Case "AU9562BSF00"
                
                Call AU9562BSF00TestSub
                
                bNeedsReStart = True
                RunReleaseMemCount = RunReleaseMemCount + 1
                
                If RunReleaseMemCount >= 10 Then
                    Call ReleaseMem(30)
                    RunReleaseMemCount = 0
                End If
            
            Case "AU9562BSF30"
                
                Call AU9562BSF30TestSub
                
                bNeedsReStart = True
                RunReleaseMemCount = RunReleaseMemCount + 1
                
                If RunReleaseMemCount >= 10 Then
                    Call ReleaseMem(30)
                    RunReleaseMemCount = 0
                End If
            
            Case "AU9562AFF00"
            
                Call AU9520ASF00TestSub
            
                bNeedsReStart = True
                RunReleaseMemCount = RunReleaseMemCount + 1
                
                If RunReleaseMemCount >= 10 Then
                    Call ReleaseMem(30)
                    RunReleaseMemCount = 0
                End If
                
            Case "AU9525CLF20"
            
                Call AU9525CLF20TestSub
                
                bNeedsReStart = True
                RunReleaseMemCount = RunReleaseMemCount + 1
                
                If RunReleaseMemCount >= 10 Then
                    Call ReleaseMem(30)
                    RunReleaseMemCount = 0
                End If
                
            Case "AU9560BSF20"
                
                Call AU9560BSF20TestSub
                
                bNeedsReStart = True
                RunReleaseMemCount = RunReleaseMemCount + 1
                
                If RunReleaseMemCount >= 10 Then
                    Call ReleaseMem(30)
                    RunReleaseMemCount = 0
                End If
                
            Case "AU9560BSF00"
            
                Call AU9560BSF00TestSub
                
                bNeedsReStart = True
                RunReleaseMemCount = RunReleaseMemCount + 1
                
                If RunReleaseMemCount >= 10 Then
                    Call ReleaseMem(30)
                    RunReleaseMemCount = 0
                End If
            
            Case "AU9540CSF20", "AU9540CSF21", "AU9562CSF20"
            
                Call SmartCard_COM_Mode_TestSub
            
            
            Case "AU661048"
           
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                End If
                
                
                 '=========================================
                '    POWER on
                '=========================================
                 CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                 
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 
                 Call MsecDelay(0.05)
                 CardResult = DO_WritePort(card, Channel_P1A, &H0)   'Power Enable
                
                 Call MsecDelay(1.2)    'power on time
                 
                  If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                 End If
           
           
                 rv0 = AU6610Test
                 
               
                 If rv0 = 1 Then
                   rv1 = 1
                   Call MsecDelay(0.2)    'power on time
                   rv2 = AU6610TestRC
                   
                  If rv2 = 0 Then  ' IR function fail
                     rv2 = 2
                   End If
                
               
                 ElseIf rv0 = 4 Then  ' speed error
        
                    rv1 = 2
                    rv2 = 4
                    
                  ElseIf rv0 = 0 Then  ' unknoew device
                  
                     rv1 = 4
                    rv2 = 4
                  
                 End If
                 
                 
                
                 
                 
                 
                 Print "rv0="; rv0
                 
                Print "rv1="; rv1
                 Print "rv2="; rv2
                Call LabelMenu(0, rv0, 1)
                Call LabelMenu(1, rv1, rv0)
                Call LabelMenu(2, rv2, rv1)
                   
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv0 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                ElseIf rv1 = WRITE_FAIL Then
                    CFWriteFail = CFWriteFail + 1
                    TestResult = "CF_WF"
                ElseIf rv1 = READ_FAIL Then
                    CFReadFail = CFReadFail + 1
                    TestResult = "CF_RF"
                ElseIf rv2 = WRITE_FAIL Then
                    XDWriteFail = XDWriteFail + 1
                    TestResult = "XD_WF"
                ElseIf rv2 = READ_FAIL Then
                    XDReadFail = XDReadFail + 1
                    TestResult = "XD_RF"
                 ElseIf rv3 = WRITE_FAIL Then
                    MSWriteFail = MSWriteFail + 1
                    TestResult = "MS_WF"
                ElseIf rv3 = READ_FAIL Then
                    MSReadFail = MSReadFail + 1
                    TestResult = "MS_RF"
                ElseIf rv0 * rv1 * rv2 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
                 
                             
                
        
        
           Case "AU6610"
           
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                End If
                
                
                 '=========================================
                '    POWER on
                '=========================================
                 CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                 
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 
                 Call MsecDelay(0.05)
                 CardResult = DO_WritePort(card, Channel_P1A, &H0)   'Power Enable
                
                 Call MsecDelay(1.2)    'power on time
                 
                  If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                 End If
           
           
                 rv0 = AU6610Test
                 
               
                 If rv0 = 1 Then
                 rv1 = 1
                  
                 ElseIf rv0 = 4 Then
        
                    rv1 = 2
                    rv0 = 1
             
                 End If
                 
                 Print "rv0="; rv0
                 
                Print "rv1="; rv1
                
                Call LabelMenu(0, rv0, 1)
                Call LabelMenu(1, rv1, rv0)
                   
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv0 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                ElseIf rv1 = WRITE_FAIL Then
                    CFWriteFail = CFWriteFail + 1
                    TestResult = "CF_WF"
                ElseIf rv1 = READ_FAIL Then
                    CFReadFail = CFReadFail + 1
                    TestResult = "CF_RF"
                ElseIf rv2 = WRITE_FAIL Then
                    XDWriteFail = XDWriteFail + 1
                    TestResult = "XD_WF"
                ElseIf rv2 = READ_FAIL Then
                    XDReadFail = XDReadFail + 1
                    TestResult = "XD_RF"
                 ElseIf rv3 = WRITE_FAIL Then
                    MSWriteFail = MSWriteFail + 1
                    TestResult = "MS_WF"
                ElseIf rv3 = READ_FAIL Then
                    MSReadFail = MSReadFail + 1
                    TestResult = "MS_RF"
                ElseIf rv0 * rv1 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
 
' ================== for 6981hlf28 ==========================
        Case "AU6981HLF28", "AU6921HLF2MF28"
            
            If ChipName = "AU6921HLF2MF28" Then
                LastCardFail = 1
                ChipName = "AU6981HLF28"
            End If
            
            If Left(ChipName, 10) = "AU6981HLF2" Then
                ChipNo = 1
            End If
 
            If LastCardFail = 1 Or FixCardMode.Value = 1 Then
                Print "Fix Card"
                If NewFixCardSub <> 1 Then
                    MPFail = True
                    GoTo End_6981HLF28
                Else
                    MPFail = False
                End If
            End If

            If PCI7248InitFinish = 0 Then
                PCI7248Exist
            End If
                
            '=========================================
            '    POWER on
            '=========================================
            
            CardResult = DO_WritePort(card, Channel_P1A, &HFF) 'ENA_B=0, ENA_A=0, SEL=1
            Call PowerSet2(1, "0", "0.2", 1, "0", "0.2", 1)
            
            If CardResult <> 0 Then
                MsgBox "Power off fail"
                End
            End If
                 
            If LastCardFail = 1 Or FixCardMode.Value = 1 Then
                Call MsecDelay(2)
                LastCardFail = 0
            Else
                Call MsecDelay(0.5)
            End If
            
            Call PowerSet2(1, "5", "0.2", 1, "5", "0.2", 1)
                  
            Call MsecDelay(0.05)
            CardResult = DO_WritePort(card, Channel_P1A, &H0)  'Power Enable
            Call MsecDelay(1.8)     'power on time
            
            LightOn = 0
            CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
            Call MsecDelay(0.5)
            
            LBA = LBA + 1   ' for dual channel
            Print LBA
               
            ClosePipe
            ReaderExist = 0
            rv0 = CBWTest_New_AU6390MB(0, 1, "vid_058f", 0)
            Call LabelMenu(0, rv0, 1)

            ClosePipe
            rv1 = CBWTest_New_AU6390MB(0, rv0, "vid_058f", 1)
            Call LabelMenu(1, rv1, rv0)

            ClosePipe
            OldLBa = LBA
            LBA = LBA + 4079614 'org= 3932160=7C7FFD*0.5
            LBA = LBA + 65535
            
            ClosePipe
            rv2 = CBWTest_New_AU6390MB(0, rv1, "vid_058f", 2)
            
            ClosePipe
            Call LabelMenu(2, rv2, rv1)
                
            '============ AU6982IL  the second 2 MBG Flash begin ==========
            If ChipName = "AU6982IL" And rv2 = 1 Then
                LBA = LBA + 4079614 'org= 3932160=7C7FFD*0.5
                LBA = LBA + 65535
                ClosePipe
                rv2 = CBWTest_New_AU6390MB(0, rv1, "vid_058f", 2)
                ClosePipe
                Call LabelMenu(2, rv2, rv1)
            End If
            
            If ChipName = "AU6982IL" And rv2 = 1 Then
                LBA = LBA + 4079614 'org= 3932160=7C7FFD*0.5
                LBA = LBA + 65535
                ClosePipe
                rv2 = CBWTest_New_AU6390MB(0, rv1, "vid_058f", 2)
                ClosePipe
                Call LabelMenu(2, rv2, rv1)
            End If
                
            '============ AU6982IL  the second 2 MBG Flash begin ==========
            LBA = OldLBa
            ClosePipe
            Call MsecDelay(0.5)
                
            If CardResult <> 0 Then
                MsgBox "Read card detect light ON fail"
                End
            End If
                   
            If ChipName = "AU6980MB" Or "AU6981HLF25" Then
                Print "Light="; LightOn
                If rv2 = 1 Then
                    If LightOn = 207 Or LightOn = 223 Then
                        rv3 = 1
                    Else
                        GPOFail = 2
                        rv3 = 2
                    End If
                Else
                    rv3 = 4
                End If
            End If
                 
            If ChipName = "AU6982IL" Then
                If rv2 = 1 Then
                    If LightOn = 254 Then
                        rv3 = 1
                    Else
                        GPOFail = 2
                        rv3 = 2
                    End If
                Else
                    rv3 = 4
                End If
            End If
 
            Call LabelMenu(32, rv3, rv2)
            CardResult = DO_WritePort(card, Channel_P1A, &HFF)
            Call PowerSet2(1, "0", "0.2", 1, "0", "0.2", 1)
                 
            If CardResult <> 0 Then
                MsgBox "Power off fail"
                End
            End If
                  
            Print "rv0=;"; rv0
            Print "rv1=;"; rv1
            Print "rv2=;"; rv2
                 
                  
            If rv0 = UNKNOW Then
                UnknowDeviceFail = UnknowDeviceFail + 1
                TestResult = "UNKNOW"
            ElseIf rv0 = WRITE_FAIL Then
                SDWriteFail = SDWriteFail + 1
                TestResult = "SD_WF"
            ElseIf rv0 = READ_FAIL Then
                SDReadFail = SDReadFail + 1
                TestResult = "SD_RF"
            ElseIf rv1 = WRITE_FAIL Then
                CFWriteFail = CFWriteFail + 1
                TestResult = "CF_WF"
            ElseIf rv1 = READ_FAIL Then
                CFReadFail = CFReadFail + 1
                TestResult = "CF_RF"
            ElseIf rv2 = WRITE_FAIL Then
                XDWriteFail = XDWriteFail + 1
                TestResult = "XD_WF"
            ElseIf rv2 = READ_FAIL Then
                XDReadFail = XDReadFail + 1
                TestResult = "XD_RF"
            ElseIf rv3 = WRITE_FAIL Then
                MSWriteFail = MSWriteFail + 1
                TestResult = "MS_WF"
            ElseIf rv3 = READ_FAIL Then
                MSReadFail = MSReadFail + 1
                TestResult = "MS_RF"
            ElseIf rv0 * rv1 * rv2 * rv3 = 1 Then
                TestResult = "PASS"
            Else
                TestResult = "Bin2"
            End If
               
            If TestResult <> "PASS" And rv0 <> UNKNOW Then
                If rv0 * rv1 * rv2 <> 1 Then
                    LastCardFail = 1
                End If
            End If
            
End_6981HLF28:
        If MPFail = True Then
            TestResult = "MP fail"
        End If
            
        ' ================== end if 6981hlf28 =======================

        ' ================== for 6981hlf30 ==========================
        Case "AU6981HLF30", "AU6981HLF3MF30"
        
            If ChipName = "AU6921HLF2MF30" Then
                LastCardFail = 1
                ChipName = "AU6981HLF30"
            End If
                      
            HVFlag = False
            LVFlag = False
                      
            If Left(ChipName, 10) = "AU6981HLF3" Then
                ChipNo = 1
            End If
 
            If LastCardFail = 1 Or FixCardMode.Value = 1 Then
                Print "Fix Card"
                If NewFixCardSub <> 1 Then
                    MPFail = True
                    GoTo End_6981HLF30
                Else
                    MPFail = False
                End If
            End If

            If PCI7248InitFinish = 0 Then
                PCI7248Exist
            End If
                
            '=========================================
            '    POWER on
            '=========================================
            
            CardResult = DO_WritePort(card, Channel_P1A, &HFF) 'ENA_B=0, ENA_A=0, SEL=1
            Call PowerSet2(1, "0", "0.2", 1, "0", "0.2", 1)
            
            If CardResult <> 0 Then
                MsgBox "Power off fail"
                End
            End If
                 
            If LastCardFail = 1 Or FixCardMode.Value = 1 Then
                Call MsecDelay(2)
                LastCardFail = 0
            Else
                Call MsecDelay(0.5)
            End If

HLV_Test:
            If HVFlag = False Then
                Call PowerSet2(1, "3.6", "0.2", 1, "3.6", "0.2", 1)
            Else
                Call PowerSet2(1, "3.0", "0.2", 1, "3.0", "0.2", 1)
            End If
                
            Call MsecDelay(0.05)
            CardResult = DO_WritePort(card, Channel_P1A, &H0)  'Power Enable
            Call MsecDelay(0.8)     'power on time
            
            LightOn = 0
            CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
            Call MsecDelay(0.5)
            
            LBA = LBA + 1   ' for dual channel
            Print LBA
               
            ClosePipe
            ReaderExist = 0
            rv0 = CBWTest_New_AU6390MB(0, 1, "vid_058f", 0)
            
            Call LabelMenu(0, rv0, 1)

            ClosePipe
            rv1 = CBWTest_New_AU6390MB(0, rv0, "vid_058f", 1)
            Call LabelMenu(1, rv1, rv0)

            ClosePipe
            OldLBa = LBA
            LBA = LBA + 4079614 'org= 3932160=7C7FFD*0.5
            LBA = LBA + 65535
            
            ClosePipe
            rv2 = CBWTest_New_AU6390MB(0, rv1, "vid_058f", 2)
            
            ClosePipe
            Call LabelMenu(2, rv2, rv1)
                
            '============ AU6982IL  the second 2 MBG Flash begin ==========
            If ChipName = "AU6982IL" And rv2 = 1 Then
                LBA = LBA + 4079614 'org= 3932160=7C7FFD*0.5
                LBA = LBA + 65535
                ClosePipe
                rv2 = CBWTest_New_AU6390MB(0, rv1, "vid_058f", 2)
                ClosePipe
                Call LabelMenu(2, rv2, rv1)
            End If
            
            If ChipName = "AU6982IL" And rv2 = 1 Then
                LBA = LBA + 4079614 'org= 3932160=7C7FFD*0.5
                LBA = LBA + 65535
                ClosePipe
                rv2 = CBWTest_New_AU6390MB(0, rv1, "vid_058f", 2)
                ClosePipe
                Call LabelMenu(2, rv2, rv1)
            End If
                
            '============ AU6982IL  the second 2 MBG Flash begin ==========
            LBA = OldLBa
            ClosePipe
            Call MsecDelay(0.5)
                
            If CardResult <> 0 Then
                MsgBox "Read card detect light ON fail"
                End
            End If
                   
            If ChipName = "AU6980MB" Or "AU6981HLF25" Then
                Print "Light="; LightOn
                If rv2 = 1 Then
                    If LightOn = 207 Or LightOn = 223 Then
                        rv3 = 1
                    Else
                        GPOFail = 2
                        rv3 = 2
                    End If
                Else
                    rv3 = 4
                End If
            End If
                 
            If ChipName = "AU6982IL" Then
                If rv2 = 1 Then
                    If LightOn = 254 Then
                        rv3 = 1
                    Else
                        GPOFail = 2
                        rv3 = 2
                    End If
                Else
                    rv3 = 4
                End If
            End If
 
            Call LabelMenu(32, rv3, rv2)
            Call PowerSet2(1, "0", "0.2", 1, "0", "0.2", 1)
            CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                 
            If CardResult <> 0 Then
                MsgBox "Power off fail"
                End
            End If
                  
            Print "rv0=;"; rv0
            Print "rv1=;"; rv1
            Print "rv2=;"; rv2
                 
            If HVFlag = False Then
            
                If rv0 * rv1 * rv2 * rv3 = 1 Then
                    HV_Result = "PASS"
                Else
                    HV_Result = "HV fail"
                End If
                
                HVFlag = True
            Else
            
                If rv0 * rv1 * rv2 * rv3 = 1 Then
                    LV_Result = "PASS"
                Else
                    LV_Result = "LV fail"
                End If
                
                LVFlag = True
                
                If HV_Result = "PASS" And LV_Result = "PASS" Then
                    TestResult = "PASS"
                ElseIf HV_Result = "HV fail" And LV_Result = "PASS" Then
                    TestResult = "BIN 3"
                ElseIf HV_Result = "PASS" And LV_Result = "LV fail" Then
                    TestResult = "BIN 4"
                ElseIf HV_Result = "HV fail" And LV_Result = "LV fail" Then
                    TestResult = "BIN 5"
                Else
                    TestResult = "BIN 2"
                End If
            
            End If
               
            If TestResult <> "PASS" And rv0 <> UNKNOW Then
                If rv0 * rv1 * rv2 <> 1 Then
                    LastCardFail = 1
                End If
            End If
            
            If LVFlag <> True Then
                GoTo HLV_Test
            End If

End_6981HLF30:
        If MPFail = True Then
            TestResult = "MP fail"
        End If
            
        ' ================== end if 6981hlf30 =======================
 
 
           Case "AU6980MB", "AU6982IL"
           
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                End If
                
                
                 '=========================================
                '    POWER on
                '=========================================
                 CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                 
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 
                 Call MsecDelay(0.05)
                 CardResult = DO_WritePort(card, Channel_P1A, &H0)  'Power Enable
                 Call MsecDelay(1.8)    'power on time
                 
                  If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                 End If
                 
           
                 
               LBA = LBA + 1   ' for dual channel
               
               Print LBA
               
                ClosePipe
                
                rv0 = CBWTest_New_AU6390MB(0, 1, "vid_058f", 0)
                 Call LabelMenu(0, rv0, 1)
                ClosePipe
                rv1 = CBWTest_New_AU6390MB(0, rv0, "vid_058f", 1)
                Call LabelMenu(1, rv1, rv0)
                 ClosePipe
                  OldLBa = LBA
                  
                  LBA = LBA + 4079614 'org= 3932160=7C7FFD*0.5
                  LBA = LBA + 65535
                 ClosePipe
                 rv2 = CBWTest_New_AU6390MB(0, rv1, "vid_058f", 2)
                 ClosePipe
                Call LabelMenu(2, rv2, rv1)
                
                '============ AU6982IL  the second 2 MBG Flash begin ==========
                If ChipName = "AU6982IL" And rv2 = 1 Then
                
                    LBA = LBA + 4079614 'org= 3932160=7C7FFD*0.5
                    LBA = LBA + 65535
                   ClosePipe
                   rv2 = CBWTest_New_AU6390MB(0, rv1, "vid_058f", 2)
                   ClosePipe
                   Call LabelMenu(2, rv2, rv1)
                End If
                
                 If ChipName = "AU6982IL" And rv2 = 1 Then
                
                    LBA = LBA + 4079614 'org= 3932160=7C7FFD*0.5
                    LBA = LBA + 65535
                    ClosePipe
                   rv2 = CBWTest_New_AU6390MB(0, rv1, "vid_058f", 2)
                   ClosePipe
                   Call LabelMenu(2, rv2, rv1)
                End If
                 '============ AU6982IL  the second 2 MBG Flash begin ==========
                
                 LBA = OldLBa
                
                ClosePipe
                
                 LightOn = 0
                 CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                  
                   If CardResult <> 0 Then
                    MsgBox "Read card detect light ON fail"
                    End
                   End If
                   
                   
                 If ChipName = "AU6980MB" Then
                       If rv2 = 1 Then
                          If LightOn = 223 Then
                             rv3 = 1
                           Else
                             GPOFail = 2
                             rv3 = 2
                          End If
                       Else
                            rv3 = 4
        
                       End If
                End If
                 
                 
                   If ChipName = "AU6982IL" Then
                       If rv2 = 1 Then
                          If LightOn = 254 Then
                             rv3 = 1
                           Else
                             GPOFail = 2
                             rv3 = 2
                          End If
                       Else
                            rv3 = 4
        
                       End If
                End If
                 
                 
                  Call LabelMenu(32, rv3, rv2)
                 
                     
                   CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                 
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                  
                  
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv0 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                ElseIf rv1 = WRITE_FAIL Then
                    CFWriteFail = CFWriteFail + 1
                    TestResult = "CF_WF"
                ElseIf rv1 = READ_FAIL Then
                    CFReadFail = CFReadFail + 1
                    TestResult = "CF_RF"
                ElseIf rv2 = WRITE_FAIL Then
                    XDWriteFail = XDWriteFail + 1
                    TestResult = "XD_WF"
                ElseIf rv2 = READ_FAIL Then
                    XDReadFail = XDReadFail + 1
                    TestResult = "XD_RF"
                 ElseIf rv3 = WRITE_FAIL Then
                    MSWriteFail = MSWriteFail + 1
                    TestResult = "MS_WF"
                ElseIf rv3 = READ_FAIL Then
                    MSReadFail = MSReadFail + 1
                    TestResult = "MS_RF"
                ElseIf rv0 * rv1 * rv2 * rv3 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
               
        
        
            Case "AU6981HLF22", "AU6982IL"
           
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                End If
                
                
                 '=========================================
                '    POWER on
                '=========================================
                 CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                 
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 
                 Call MsecDelay(0.05)
                 CardResult = DO_WritePort(card, Channel_P1A, &H0)  'Power Enable
                 Call MsecDelay(1.8)    'power on time
                 
                  If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                 End If
                 
      
               
                ClosePipe
                 
                rv0 = CBWTest_New_Physical_Read(0, 1, "vid_058f")
                 Call LabelMenu(0, rv0, 1)
                ClosePipe
                
                Call LabelMenu(2, rv2, rv1)
  
                
                 CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                  
                   If CardResult <> 0 Then
                    MsgBox "Read card detect light ON fail"
                    End
                   End If
                   
                   
               
                       If rv0 = 1 Then
                          If LightOn = 223 Then
                             rv3 = 1
                           Else
                             GPOFail = 2
                             rv3 = 2
                          End If
                       Else
                            rv3 = 4
        
                       End If
              
                 Call LabelMenu(2, rv3, rv0)
                 
                 
                 
                 
                     
                   CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                 
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                  
                  
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv0 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                ElseIf rv1 = WRITE_FAIL Then
                    CFWriteFail = CFWriteFail + 1
                    TestResult = "CF_WF"
                ElseIf rv1 = READ_FAIL Then
                    CFReadFail = CFReadFail + 1
                    TestResult = "CF_RF"
                ElseIf rv2 = WRITE_FAIL Then
                    XDWriteFail = XDWriteFail + 1
                    TestResult = "XD_WF"
                ElseIf rv2 = READ_FAIL Then
                    XDReadFail = XDReadFail + 1
                    TestResult = "XD_RF"
                 ElseIf rv3 = WRITE_FAIL Then
                    MSWriteFail = MSWriteFail + 1
                    TestResult = "MS_WF"
                ElseIf rv3 = READ_FAIL Then
                    MSReadFail = MSReadFail + 1
                    TestResult = "MS_RF"
                ElseIf rv0 * rv3 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
     
     
                
            
            Case "AU9520JJ"
            
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                End If
                
                
                 '=========================================
                '    POWER on
                '=========================================
                 CardResult = DO_WritePort(card, Channel_P1CL, &HF)
                 
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 
                 Call MsecDelay(0.05)
                 CardResult = DO_WritePort(card, Channel_P1CL, &HB)  'Power Enable
                 Call MsecDelay(0.8)    'power on time
                 
                  If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                 End If
                 
                 
                 '===========================================
                 'NO card test
                 '============================================
                   CardResult = DO_ReadPort(card, Channel_P1CH, LightOff)
                  
                   If CardResult <> 0 Then
                    MsgBox "Read card detect light off fail"
                    End
                   End If
                   
                  
                   
                 
                '===========================================
                 'NO card test
                 '============================================
                  CardResult = DO_WritePort(card, Channel_P1CL, &HA)  'Power Enable + Slot0 enable
                   If CardResult <> 0 Then
                    MsgBox "Set card detect light on fail"
                    End
                   End If
                   Call MsecDelay(0.05)
                    CardResult = DO_ReadPort(card, Channel_P1CH, LightOn)
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light on fail"
                    End
                   End If
                   
                   
                  
                '===========================================
                 'R/W test
                 '============================================
                  
                If LightOff <> 14 Or LightOn <> 12 Then
                  TestResult = "Bin2"
                Else
            
                        PreviousStatus = 0
                        TestResultadd = "FAIL"
                        TestResultSmartCard = "FAIL" 'arch add 0321
                        ArchTest = False             'arch add 0321
                        Cls
                        
                        strReaderName = VenderString_AU9520JJ & " " & ProductString_AU9520JJ & " 0"   'slot 0
                        udtReaderStates(0).szReader = strReaderName & vbNullChar
                        txtmsg.Text = udtReaderStates(0).szReader
                        Print "Start to test slot0!"
                        Call StartTest
                        
                        If TestResultadd = "PASS" Then 'Arch define 940410
                       ' Call LabelMenu(5, 1, 1)
                            TestResult = "PASS"
                        Else
                            'TestResult = "FAIL"
                            TestResult = "Bin2"
                        End If
                End If
        
        
         Case "AU6366ALF20", "AU6366CLF20", "AU6371GLF20", "AU6371GLF24", "AU6371ELF24", "AU6371DLF24", "AU6371HLF24", "AU6433EFF21", "AU6433DFF21", "AU6433HFF21", "AU6433KFF21", "AU6433BSF20", "AU6433GSF21"
     
              
               Call SingleSlotTest
               
        Case "AU8451DBF22", "AU8451BBF22", "AU8451EBF22"
        
            Call AU8451FTTestSub
               
        Case "AU6485AFF25", "AU6485AFF05", "AU6485BFF25", "AU6485CFF05", "AU6485CFF25", "AU6485DFF25", "AU6485HFF05", "AU6485HFF25", "AU6485AFS15", "AU6485CFS15", "AU6485HFS15", "AU6485IFF25", "AU6485JFF25", "AU6485LFF25"
        
            Call AU6485FTTestSub
        
        Case "AU6435DLF20", "AU6435DLF21", "AU6435DLF22", "AU6435DLF23", "AU6435DLF24", "AU6435ELF21", "AU6435ELF22", "AU6435ELF23", "AU6435ELF24", "AU6435ELF25", "AU6435GLE10", "AU6435ELF33", "AU6435ELF34", "AU6435ELS11", "AU6435DLF2D", "AU6435BLF21", "AU6435AFF21", "AU6435BFF20", "AU6435CFF22", "AU6435BFF21", "AU6435BFF23", "AU6435BFS11"
               
            Call AU6435TestSub
        
        Case "AU6435ELF04"
        
            Call AU6435TestSub
        
        Case "AU6465CFF20"
        
            Call AU6465TestSub

        Case "AU6479ALF20", "AU6479ALT10", "AU6479AFF33", "AU6479BLF20", "AU6479BLF21", "AU6479HLF21", "AU6479BLF23", "AU6479JLF23", "AU6479NLF23", "AU6479ILF23", "AU6479HLF22", "AU6479KLF23", "AU6479KLF24", "AU6479FLF23", "AU6479FLF03", "AU6479OLF23", "AU6479TLF23", "AU6479BFF23", "AU6479OLT10", "AU6479OLF03", "AU6479CFF23", "AU6479DFF24", "AU6479ULF23", "AU6479FLF24", "AU6479ULF24", "AU6479WLF22", "AU6479PLF32"
        
            Call AU6479TestSub
        
        Case "AU6479BLF02", "AU6479JLFE3"
        
            Call AU6479TestSub
        
         Case "AU6433LFF22", "AU6433IFF22", "AU6433HSF22", "AU6433FSF22", "AU6433ESF22", "AU6433EFF22", "AU6433KFF22", "AU6433GSF22", "AU6433DFF22", "AU6433HFF22"
         
         
               Call AU6433F22MPTest
        
        Case "AU6433VFF23"
        
            Call AU6433F23MPTest
                
        Case "AU6433FSF29", "AU6433FSF28", "AU6433CSF29", "AU6433CSF2A", "AU6433CSF09", "AU6433CSF0A", "AU6433CSF28", "AU6433DLF20", "AU6433DLF21", "AU6433DLF22", "AU6433DLF03", "AU6433DLF23", "AU6433DLF30", "AU6433DLF31", "AU6433DLF32", "AU6433DLF33", "AU6433DLF3C", "AU6433JSF29", "AU6433JSF39", "AU6433JSF28", "AU6433BLF2A", "AU6433BLF2A", "AU6433BLF3B", "AU6433BLF3C", "AU6433BLD3B", "AU6433BLF29", "AU6433BLF2C", "AU6433BLF30", "AU6433BLF28", "AU6433BLF27", "AU6433BLS13", "AU6433BLS12", "AU6433BLS11", "AU6433BLF26", "AU6433BLS10", "AU6433EFF35", "AU6433EFF36", "AU6433EFF3F", "AU6433EFF25", "AU6433BLF24", "AU6433LFF33", "AU6433LFF23", "AU6433IFF23", "AU6433HSF23", "AU6433FSF23", "AU6433ESF23", "AU6433EFF23", "AU6433KFF23", "AU6433GSF23", "AU6433DFF23", "AU6433HFF23"
         
               Call AU6433F23MPTest
               
            Case "AU6433BLF3D", "AU6433BLF3E", "AU6433BLF3F", "AU6433BLF2E", "AU6433BLF2F", "AU6433BLFEE", "AU6433BLF0E", "AU6433BLF0F", "AU6433BLF2D", "AU6433DLF00", "AU6433EFF26"
            
                Call AU6433F23MPTest

            
            Case "AU6433JSF2A", "AU6433JSF3A", "AU6433FSF2A", "AU6433FSF0A"
            
                Call AU6433F2AMPTest
            
            Case "AU6433BLS2F", "AU6433BLS0F"
            
                Call AU6433S61MPTest
            
           'Case "AU6438CFF20", "AU6438BSF20", "AU6438BSF21", "AU6438BSF00", "AU6438BSF01", "AU6438BSF02", "AU6438EFS20", "AU6438GFF20", "AU6438BLF20", "AU6438CLF20", "AU6438IFF20", "AU6438IFS30", "AU6438KFS00", "AU6438KFS30", "AU6438IFS00", "AU6438CFF00", "AU6438CFS10", "AU6438KFE10", "AU6438KFS10", "AU6438MFS20", "AU6438BSS21", "AU6438BSS01", "AU6438CFS21"
            Case "AU6438BSS21", "AU6438BSS01", "AU6438CFS21", "AU6438EFS20", "AU6438MFS20", "AU6438IFS30", "AU6438IFS00", "AU6438KFS30", "AU6438KFS00" _
               , "AU6438BSF21", "AU6438BSF01", "AU6438CFF01", "AU6438CFF21", "AU6438EFF20", "AU6438MFF20", "AU6438IFF30", "AU6438IFF00", "AU6438KFF30", "AU6438KFF00"
                                
            
               Call AU6438MPTest
               
            Case "AU9540DSF22"
                
                Call AU9540TestSub
                
            Case "AU9540DSF02"
                
                Call AU9540Test_HLV_Sub
                
          Case "AU6371DLF25", "AU6371ELF25", "AU6371PLF25"
                 
                 If ChipName = "AU6371PLF25" Then
                 
                   ChipName = "AU6371DLF25"
                 End If
                 
                 Call SingleSlotTest25
                 
           Case "AU6371ELF27", "AU6371ELS30", "AU6371DLF27", "AU6371TLF27", "AU6371SLF27", "AU6371HLF27", "AU6371GLF27"
                 
                    Call SingleSlotTest27
                    
            Case "AU6371ELF28", "AU6371DLF28", "AU6371TLF28", "AU6371SLF28", "AU6371HLF28", "AU6371GLF28"
            
                    Call AU6371MP28
            
            
             Case "AU6371DLS60", "AU6371DLS50", "AU6371DLF2B", "AU6371DLF2A", "AU6371DLO10", "AU6371ELF29", "AU6371DLF29", "AU6371TLF29", "AU6371SLF29", "AU6371HLF29", "AU6371DLF09", "AU6371GLF29"
            
                    If ChipName = "AU6371DLF2A" Then
                        AU6371DLF2ATestSub
                    ElseIf ChipName = "AU6371DLF2B" Then
                        AU6371DLF2BTestSub
                    ElseIf ChipName = "AU6371DLS50" Then
                        AU6371DLS50SortingSub
                    ElseIf ChipName = "AU6371DLS60" Then
                        AU6371DLS60SortingSub
                    ElseIf ChipName = "AU6371DLF09" Then
                        AU6371DLF09TestSub
                    Else
                        Call AU6371MP29
                    End If
            
                   
            Case "AU6371DLS22", "AU6371DLF26", "AU6371ELF26", "AU6371PLF26", "AU6371NLF26", "AU6371SLF26", "AU6371TLF26", "AU6371DLS10", "AU6371DLS20", "AU6371DLS30", "AU6371DLS31", "AU6371DLS21", "AU6371DLS32", "AU6371DLS40", "AU6371DLS41"
                 
                 If ChipName = "AU6371PLF26" Then
                 
                   ChipName = "AU6371DLF26"
                 End If
                 
                 Call SingleSlotTest26
                
            Case "AU6471GLF21", "AU6471FLF21"
            
                   Call AU6471F21TestSub
                   
                   
              Case "AU6471JLS10", "AU6471GLF23", "AU6471GLF24", "AU6471GLF04", "AU6471FLF23", "AU6471GLF22", "AU6471FLF22"
            
                   Call AU6471F22TestSub
                  
                  
            Case "AU6471FLF20", "AU6471GLF20"
                 
                  Call SingleSlotTest26
                       
            Case "AU6256XLSE2", "AU6256XLS32", "AU6256XLS31", "AU6256XLS21", "AU6256XLS20", "AU6256XLS1D", "AU6256XLS1C", "AU6256XLS1B", "AU6256XLS1A", "AU6256XLS19", "AU6256XLS18", "AU6256XLS17", "AU6256XLS16", "AU6256XLS15", "AU6256XLS14", "AU6256XLS13", "AU6256XLS10", "AU6256XLS11", "AU6256XLS12", "AU6471FLS11", "AU6337BLS10", "AU698XHLS20", "AU6430BLS10", "AU6476BLS10", "AU6433EFS10", "AU6433DFS10", "AU6471FLS10", "AU6430BLF22", "AU6430DLF22", "AU6430ELF22", "AU6430QLF22", "AU698XHLS10", "AU6430QLF21", "AU6430QLF20", "AU6430QLS10", "AU6430ELS10", "AU6430DLS10", "AU6430QLS11", "AU6430ELS11", "AU6430DLS11", "AU6430QLS12", "AU6430ELS12", "AU6430DLS12", _
                   "AU6430QLS13", "AU6430ELS13", "AU6430DLS13", "AU6430DLF20", "AU6430ELF20", "AU6430BLF20", "AU6430ELS20"
            
               Call AU6430MPSub
                 
            Case "AU6473CLF20", "AU6473CLF21", "AU6473CLF31", "AU6473BLF20", "AU6429FLF20", "AU6425DLF20", "AU6427ELF20", "AU6427GLF20"
                
                Call AU6473TestSub
                 
              Case "AU6476MLF24", "AU6476FLF21", "AU6476DLF22", "AU6476CLF22", "AU6476ELF21", "AU6476BLF23", "AU6476ILF22", "AU6476BLF22", "AU6420ALF20", "AU6420BLF20", "AU6420BLF2A", "AU6420BLS2A", "AU6420CLF20", "AU6476BLF20", "AU6476CLF20", "AU6476CLF21", "AU6476FLF20", "AU6476DLF20", "AU6476ELF20", "AU6476ILF21", "AU6476BLF21", "AU6476DLF21"
            
                
                    Call AU6476TestSub
                   
             Case "AU6420BLS20", "AU6420BLS30"
             
                Call AU6476TestSub
                    
             Case "AU6476RLF27", "AU6476RLF26", "AU6476RLF25", "AU6476JLO20", "AU6476QLF25", "AU6476LLF25", "AU6476JLOT6", "AU6476JLOT1", "AU6476JLOT2", "AU6476JLOT3", "AU6476JLO10", "AU6476BLF24", "AU6476BLF25", "AU6476WLF05", "AU6476WLF35", "AU6476WLF3E", "AU6476CLF26", "AU6476CLF06", "AU6476CLF25", "AU6476CLF24", "AU6476DLF24", "AU6476ELF24", "AU6476FLF24", "AU6476ILF24", "AU6476KLF20"
             
                Call AU6476TestSub
                
            Case "AU6476CLF27", "AU6476CLF07", "AU6476VLF26", "AU6476YLF26"
            
                Call AU6476_020222TestSub
            
                  
            Case "AU6476BLF26", "AU6476WLF06", "AU6476WLF36", "AU6476WLF3F"
            
                Call AU6476TestSub
            
            Case "AU6352LLF20", "AU6352DFF20", "AU6352LLF00"
                
                Call AU6352TestSub
            
           Case "AU6371ELS10"
           
           
                Call AU6371ELS10Normal
                
                  
           Case "AU6371ELS11"
           
           
                Call AU6371ELS11Ram
                
                 
        Case "AU6371ELS20", "AU6980HLS20"
                If ChipName = "AU6371ELS20" Then
                 Call AU6371ELS20MSPro
               ElseIf ChipName = "AU6980HLS20" Then
                 Call AU6980HLS20SortingSub
               End If
                 
                
            Case "AU6375F", "AU6368F", "AU6376HV"
            
                
                If ChipName = "AU6376HV" Then
                         Call PowerSet(21)
                         If PCI7248InitFinish = 0 Then
                          PCI7248Exist
                         End If
                         
                         CardResult = DO_WritePort(card, Channel_P1B, &H0)
                        result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                         CardResult = DO_WritePort(card, Channel_P1A, &H0)  '
                           
                         Call MsecDelay(1.2)
                          
               End If
                
                
                
                
                rv0 = 0
                rv1 = 0
                rv2 = 0
                rv3 = 0
                rv4 = 0
                rv5 = 0
                rv6 = 0
                rv7 = 0
                
                LBA = LBA + 1
                
                Label3.BackColor = RGB(255, 255, 255)
                Label4.BackColor = RGB(255, 255, 255)
                Label5.BackColor = RGB(255, 255, 255)
                Label6.BackColor = RGB(255, 255, 255)
                Label7.BackColor = RGB(255, 255, 255)
                Label8.BackColor = RGB(255, 255, 255)
                   
            
                OldLBa = LBA
                ClosePipe
                rv0 = CBWTest_New(0, 1, "vid_058f")
                Call LabelMenu(0, rv0, 1)
                '//*LabelMenu(SlotNo As Byte, TestResult As Byte, PreSlotStatus As Byte)
                ClosePipe
                rv1 = CBWTest_New(1, rv0, "vid_058f")
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
                rv2 = CBWTest_New(2, rv1, "vid_058f")
                Call LabelMenu(2, rv2, rv1)
                ClosePipe
                rv3 = CBWTest_New(3, rv2, "vid_058f")
                Call LabelMenu(3, rv3, rv2)
                ClosePipe
                rv4 = CBWTest_New_8_Sector(0, rv3, FailPosition, 20)
                LBA = OldLBa
                Call LabelMenu(0, rv4, rv3)
                ClosePipe
                rv5 = CBWTest_New_8_Sector(1, rv4, FailPosition, 20)
                LBA = OldLBa
                Call LabelMenu(1, rv5, rv4)
                ClosePipe
                rv6 = CBWTest_New_8_Sector(2, rv5, FailPosition, 20)
                LBA = OldLBa
                Call LabelMenu(2, rv6, rv5)
                ClosePipe
                rv7 = CBWTest_New_8_Sector(3, rv6, FailPosition, 20)
                LBA = OldLBa
                Call LabelMenu(3, rv7, rv6)
                ClosePipe
                
                
                Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print rv3, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print rv4, " \\SD stress :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print rv5, " \\CF stress :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print rv6, " \\XD stress :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print rv7, " \\MS stress :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print "LBA="; LBA
                Print "FailPosition="; FailPosition
                
                
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Or rv4 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv0 = READ_FAIL Or rv4 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                ElseIf rv1 = WRITE_FAIL Or rv5 = WRITE_FAIL Then
                    CFWriteFail = CFWriteFail + 1
                    TestResult = "CF_WF"
                ElseIf rv1 = READ_FAIL Or rv5 = READ_FAIL Then
                    CFReadFail = CFReadFail + 1
                    TestResult = "CF_RF"
                ElseIf rv2 = WRITE_FAIL Or rv6 = WRITE_FAIL Then
                    XDWriteFail = XDWriteFail + 1
                    TestResult = "XD_WF"
                ElseIf rv2 = READ_FAIL Or rv6 = READ_FAIL Then
                    XDReadFail = XDReadFail + 1
                    TestResult = "XD_RF"
                ElseIf rv3 = WRITE_FAIL Or rv7 = WRITE_FAIL Then
                    MSWriteFail = MSWriteFail + 1
                    TestResult = "MS_WF"
                ElseIf rv3 = READ_FAIL Or rv7 = READ_FAIL Then
                    MSReadFail = MSReadFail + 1
                    TestResult = "MS_RF"
                ElseIf rv7 * rv6 * rv5 * rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                End If
                
                Call PowerSet(3)
                 CardResult = DO_WritePort(card, Channel_P1A, &H1)
                  result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
            
                   CardResult = DO_WritePort(card, Channel_P1B, &H0)
                   
           
            
            Case "AU6369CF"
            
                rv0 = 0
                rv1 = 0
                
                LBA = LBA + 1
                ClosePipe
                rv0 = CBWTest_New(0, 1, "vid_058f")
              
                Call LabelMenu(1, rv0, 1)
                ClosePipe
                               
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = PASS Then 'Arch define 941027
                   TestResult = "PASS"
                ElseIf rv0 = WRITE_FAIL Then
                    CFWriteFail = CFWriteFail + 1
                    TestResult = "CF_WF"
                ElseIf rv0 = READ_FAIL Then
                    CFReadFail = CFReadFail + 1
                    TestResult = "CF_RF"
                Else
                    TestResult = "Bin2"
                End If
                        
                                
                Print rv0, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print "LBA="; LBA
                
           
        
            Case "AU6368NC"
                  
                     
              
                LBA = LBA + 1
                
                
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
                
                
                Print LBA
                
                ClosePipe
                rv0 = CBWTest_New_no_card(0, 1, "vid_058f")
                'Print "a1"
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                rv1 = CBWTest_New_no_card(1, rv0, "vid_058f")
               '  Print "a2"
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
                
                rv2 = CBWTest_New_no_card(2, rv1, "vid_058f")
               '  Print "a3"
                Call LabelMenu(2, rv2, rv1)
                ClosePipe
                rv3 = CBWTest_New_no_card(3, rv2, "vid_058f")
                ' Print "a4"
                ClosePipe
                Call LabelMenu(3, rv3, rv2)
                
                
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv0 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                ElseIf rv1 = WRITE_FAIL Then
                    CFWriteFail = CFWriteFail + 1
                    TestResult = "CF_WF"
                ElseIf rv1 = READ_FAIL Then
                    CFReadFail = CFReadFail + 1
                    TestResult = "CF_RF"
                ElseIf rv2 = WRITE_FAIL Then
                    XDWriteFail = XDWriteFail + 1
                    TestResult = "XD_WF"
                ElseIf rv2 = READ_FAIL Then
                    XDReadFail = XDReadFail + 1
                    TestResult = "XD_RF"
                 ElseIf rv3 = WRITE_FAIL Then
                    MSWriteFail = MSWriteFail + 1
                    TestResult = "MS_WF"
                ElseIf rv3 = READ_FAIL Then
                    MSReadFail = MSReadFail + 1
                    TestResult = "MS_RF"
                ElseIf rv3 * rv2 * rv1 * rv0 = PASS Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
                
                Print "Test Result"; TestResult
                       
               
              
                 MSComm1.OutBufferCount = 0
                 MSComm1.Output = TestResult   ' send out test result
                 
                 
                Print rv0, " \\SD :0 Unknow device, 1 pass ,2 card change bit fail"
                Print rv1, " \\CF :0 Unknow device, 1 pass ,2 card change bit fail"
                Print rv2, " \\XD :0 Unknow device, 1 pass ,2 card change bit fail"
                Print rv3, " \\MS :0 Unknow device, 1 pass ,2 card change bit fail"
                 
                 
                If TestResult = "PASS" Then
                  
                     TestResult = ""
                      ChipName = ""
                        Do
                            Call MsecDelay(0.1)
                            DoEvents
                            buf = MSComm1.Input
                            ChipName = ChipName & buf
                        Loop Until InStr(1, ChipName, "PASS") <> 0
                 
                 
                 '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                 '
                 '  R/W test
                 '
                 '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                 
                
                'initial return value
                
                
                rv0 = 0
                rv1 = 0
                rv2 = 0
                rv3 = 0
               
                Label3.BackColor = RGB(255, 255, 255)
                Label4.BackColor = RGB(255, 255, 255)
                Label5.BackColor = RGB(255, 255, 255)
                Label6.BackColor = RGB(255, 255, 255)
                Label7.BackColor = RGB(255, 255, 255)
                Label8.BackColor = RGB(255, 255, 255)
                
                ClosePipe
                rv0 = CBWTest_New(0, 1, "vid_058f")
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                rv1 = CBWTest_New(1, rv0, "vid_058f")
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
                rv2 = CBWTest_New(2, rv1, "vid_058f")
                Call LabelMenu(2, rv2, rv1)
                ClosePipe
                rv3 = CBWTest_New(3, rv2, "vid_058f")
                Call LabelMenu(3, rv3, rv2)
                ClosePipe
                
                
                Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print rv3, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print "LBA="; LBA
                
                
                'If rv0 = 1 And rv1 = 1 And rv2 = 1 And rv3 = 1 Then
                   ' TestResult = "PASS"
                'End If
                
                        
                        If rv0 = UNKNOW Then
                           UnknowDeviceFail = UnknowDeviceFail + 1
                           TestResult = "UNKNOW"
                        ElseIf rv0 = WRITE_FAIL Then
                            SDWriteFail = SDWriteFail + 1
                            TestResult = "SD_WF"
                        ElseIf rv0 = READ_FAIL Then
                            SDReadFail = SDReadFail + 1
                            TestResult = "SD_RF"
                        ElseIf rv1 = WRITE_FAIL Then
                            CFWriteFail = CFWriteFail + 1
                            TestResult = "CF_WF"
                        ElseIf rv1 = READ_FAIL Then
                            CFReadFail = CFReadFail + 1
                            TestResult = "CF_RF"
                        ElseIf rv2 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                        ElseIf rv2 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                         ElseIf rv3 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv3 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                        ElseIf rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
                
               ' If rv2 * rv3 * rv1 = 0 Then
                '    MsgBox "rturn error"
                'End If
                  
                End If
           
            Case "AU6331"
                LBA = LBA + 1
                
                
                TestResult = ""
                ChipName = ""
                
                rv0 = 0
                rv1 = 0
                rv2 = 0
                rv3 = 0
                
                Label3.BackColor = RGB(255, 255, 255)
                Label4.BackColor = RGB(255, 255, 255)
                Label5.BackColor = RGB(255, 255, 255)
                Label6.BackColor = RGB(255, 255, 255)
                Label7.BackColor = RGB(255, 255, 255)
                Label8.BackColor = RGB(255, 255, 255)
                
                ClosePipe
                rv0 = CBWTest_New(0, 1, "vid_058f")
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                rv1 = CBWTest_New(1, rv0, "vid_058f")
                Call LabelMenu(3, rv1, rv0)
                ClosePipe
                
                Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print rv1, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print "LBA="; LBA
                
                        
                    If rv0 = UNKNOW Then
                       UnknowDeviceFail = UnknowDeviceFail + 1
                       TestResult = "UNKNOW"
                    ElseIf rv0 = WRITE_FAIL Then
                        SDWriteFail = SDWriteFail + 1
                        TestResult = "SD_WF"
                    ElseIf rv0 = READ_FAIL Then
                        SDReadFail = SDReadFail + 1
                        TestResult = "SD_RF"
                    ElseIf rv1 = WRITE_FAIL Then
                        MSWriteFail = MSWriteFail + 1
                        TestResult = "MS_WF"
                    ElseIf rv1 = READ_FAIL Then
                        MSReadFail = MSReadFail + 1
                        TestResult = "MS_RF"
                    ElseIf rv1 * rv0 = PASS Then
                         TestResult = "PASS"
                    Else
                        TestResult = "Bin2"
                    End If
                        
          Case "AU6367"
                LBA = LBA + 1
                
                
                TestResult = ""
                ChipName = ""
                 
                rv0 = 0
                rv1 = 0
                rv2 = 0
                rv3 = 0
                
                Label3.BackColor = RGB(255, 255, 255)
                Label4.BackColor = RGB(255, 255, 255)
                Label5.BackColor = RGB(255, 255, 255)
                Label6.BackColor = RGB(255, 255, 255)
                Label7.BackColor = RGB(255, 255, 255)
                Label8.BackColor = RGB(255, 255, 255)
                
                ClosePipe
                rv0 = CBWTest_New(0, 1, "vid_058f")
                Call LabelMenu(2, rv0, 1)
                ClosePipe
                rv1 = CBWTest_New(1, rv0, "vid_058f")
                Call LabelMenu(0, rv1, rv0)
                ClosePipe
               
                Print rv0, " \\UFD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print rv1, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print "LBA="; LBA
                   
                    If rv0 = UNKNOW Then
                       UnknowDeviceFail = UnknowDeviceFail + 1
                       TestResult = "UNKNOW"
                    ElseIf rv0 = WRITE_FAIL Then
                        XDWriteFail = XDWriteFail + 1
                        TestResult = "XD_WF"
                    ElseIf rv0 = READ_FAIL Then
                        XDReadFail = XDReadFail + 1
                        TestResult = "XD_RF"
                    ElseIf rv1 = WRITE_FAIL Then
                        SDWriteFail = SDWriteFail + 1
                        TestResult = "SD_WF"
                    ElseIf rv1 = READ_FAIL Then
                        SDReadFail = SDReadFail + 1
                        TestResult = "SD_RF"
                    ElseIf rv1 * rv0 = PASS Then
                         TestResult = "PASS"
                    Else
                        TestResult = "Bin2"
                    End If
                    
            Case "AU6367F"
             
                LBA = LBA + 1
                TestResult = ""
                ChipName = ""
                 
                rv0 = 0
                rv1 = 0
                rv4 = 0
                rv5 = 0
                
                Label3.BackColor = RGB(255, 255, 255)
                Label4.BackColor = RGB(255, 255, 255)
                Label5.BackColor = RGB(255, 255, 255)
                Label6.BackColor = RGB(255, 255, 255)
                Label7.BackColor = RGB(255, 255, 255)
                Label8.BackColor = RGB(255, 255, 255)
                
                ClosePipe
                rv0 = CBWTest_New(0, 1, "vid_058f")
                Call LabelMenu(2, rv0, 1)
                ClosePipe
                rv1 = CBWTest_New(1, rv0, "vid_058f")
                Call LabelMenu(0, rv1, rv0)
                ClosePipe
                rv4 = CBWTest_New_8_Sector(0, rv1, FailPosition, 20)
                LBA = OldLBa
                Call LabelMenu(2, rv4, rv1)
                ClosePipe
                rv5 = CBWTest_New_8_Sector(1, rv4, FailPosition, 20)
                LBA = OldLBa
                Call LabelMenu(0, rv5, rv4)
                ClosePipe
                
                
                Print rv0, " \\UFD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print rv1, " \\SD  :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print rv4, " \\UFD Stress :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print rv5, " \\SD  Stress :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print "LBA="; LBA
                Print "FailPosition="; FailPosition
              
                   
                    If rv0 = UNKNOW Then
                       UnknowDeviceFail = UnknowDeviceFail + 1
                       TestResult = "UNKNOW"
                    ElseIf rv0 = WRITE_FAIL Or rv4 = WRITE_FAIL Then
                        XDWriteFail = XDWriteFail + 1
                        TestResult = "XD_WF"
                    ElseIf rv0 = READ_FAIL Or rv4 = READ_FAIL Then
                        XDReadFail = XDReadFail + 1
                        TestResult = "XD_RF"
                    ElseIf rv1 = WRITE_FAIL Or rv5 = WRITE_FAIL Then
                        SDWriteFail = SDWriteFail + 1
                        TestResult = "SD_WF"
                    ElseIf rv1 = READ_FAIL Or rv5 = READ_FAIL Then
                        SDReadFail = SDReadFail + 1
                        TestResult = "SD_RF"
                    ElseIf rv0 * rv1 * rv4 * rv5 = PASS Then
                         TestResult = "PASS"
                    Else
                        TestResult = "Bin2"
                    End If
                    
        Case "AU33NC_F"
            
                LBA = LBA + 1
                OldLBa = LBA
                
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
            
            
            Print LBA
            
                ClosePipe
                rv0 = CBWTest_New_no_card(0, 1, "vid_058f")
                'Print "a1"
            Call LabelMenu(0, rv0, 1)
                ClosePipe
                rv1 = CBWTest_New_no_card(1, rv0, "vid_058f")
                '  Print "a2"
            Call LabelMenu(3, rv1, rv0)
                ClosePipe
            
            
            
            If rv0 = UNKNOW Then
                    UnknowDeviceFail = UnknowDeviceFail + 1
                    TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv0 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                ElseIf rv1 = WRITE_FAIL Then
                    MSWriteFail = MSWriteFail + 1
                    TestResult = "MS_WF"
                ElseIf rv1 = READ_FAIL Then
                    MSReadFail = MSReadFail + 1
                    TestResult = "MS_RF"
                ElseIf rv1 * rv0 = PASS Then
                    TestResult = "PASS"
                Else
                    TestResult = "Bin2"
            End If
            
            Print "Test Result"; TestResult
            
            MSComm1.OutBufferCount = 0
            
            MSComm1.Output = TestResult   ' send out test result
            
            
            Print rv0, " \\SD :0 Unknow device, 1 pass ,2 card change bit fail"
            Print rv1, " \\MS :0 Unknow device, 1 pass ,2 card change bit fail"
            
            
            If TestResult = "PASS" Then
            
                TestResult = ""
                ChipName = ""
                Do
                
                    Call MsecDelay(0.1)
                    DoEvents
                    buf = MSComm1.Input
                    ChipName = ChipName & buf
                
                Loop Until InStr(1, ChipName, "PASS") <> 0
            
            
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '
                '  R/W test
                '
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                
                
                'initial return value
                
                
                rv0 = 0
                rv1 = 0
                rv2 = 0
                rv3 = 0
                rv4 = 0
                rv5 = 0
                
                Label3.BackColor = RGB(255, 255, 255)
                Label4.BackColor = RGB(255, 255, 255)
                Label5.BackColor = RGB(255, 255, 255)
                Label6.BackColor = RGB(255, 255, 255)
                Label7.BackColor = RGB(255, 255, 255)
                Label8.BackColor = RGB(255, 255, 255)
                
                
                
                ClosePipe
                rv0 = CBWTest_New(0, 1, "vid_058f")
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                rv1 = CBWTest_New(1, rv0, "vid_058f")
                Call LabelMenu(3, rv1, rv0)
                ClosePipe
                rv4 = CBWTest_New_8_Sector(0, rv1, FailPosition, 20)
                LBA = OldLBa
                Call LabelMenu(0, rv4, rv1)
                ClosePipe
                rv5 = CBWTest_New_8_Sector(1, rv4, FailPosition, 20)
                LBA = OldLBa
                Call LabelMenu(1, rv5, rv4)
                ClosePipe
                
                
                Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print rv1, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print rv4, " \\SD Stress :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print rv5, " \\MS Stress:0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print "LBA="; LBA
                Print "FailPosition="; FailPosition
                
                
                If rv0 = UNKNOW Then
                    UnknowDeviceFail = UnknowDeviceFail + 1
                    TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Or rv4 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv0 = READ_FAIL Or rv4 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                ElseIf rv1 = WRITE_FAIL Or rv5 = WRITE_FAIL Then
                    CFWriteFail = CFWriteFail + 1
                    TestResult = "CF_WF"
                ElseIf rv1 = READ_FAIL Or rv5 = READ_FAIL Then
                    CFReadFail = CFReadFail + 1
                    TestResult = "CF_RF"
                ElseIf rv2 = WRITE_FAIL Or rv6 = WRITE_FAIL Then
                    XDWriteFail = XDWriteFail + 1
                    TestResult = "XD_WF"
                ElseIf rv2 = READ_FAIL Or rv6 = READ_FAIL Then
                    XDReadFail = XDReadFail + 1
                    TestResult = "XD_RF"
                ElseIf rv3 = WRITE_FAIL Or rv7 = WRITE_FAIL Then
                    MSWriteFail = MSWriteFail + 1
                    TestResult = "MS_WF"
                ElseIf rv3 = READ_FAIL Or rv7 = READ_FAIL Then
                    MSReadFail = MSReadFail + 1
                    TestResult = "MS_RF"
                ElseIf rv5 * rv4 * rv1 * rv0 = PASS Then
                    TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                End If
            
            End If
            
          
                    
                    
          Case "AU6331A1"
         
         
                'GPIO control setting
            
                 
                    If PCI7248InitFinish = 0 Then
                      PCI7248Exist
                    End If
                    
                    
                
                     
                    CardResult = DO_WritePort(card, Channel_P1A, &HE)
                    Call MsecDelay(1#)
                   
                 
            
                  
                rv0 = 0
                rv1 = 0
                rv2 = 0
                rv3 = 0
                rv4 = 0
                rv5 = 0
                rv6 = 0
                rv7 = 0
                
                LBA = LBA + 1
                
                
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
                
                
                Print LBA
                
                ClosePipe
                rv0 = CBWTest_New_no_card(0, 1, "vid_058f")
                 
                Call LabelMenu(0, rv0, 1)
              
                ClosePipe
                
                
                
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv0 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
               
                ElseIf rv0 = PASS Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                End If
                
                Print "Test Result"; TestResult
                       
             
                 
                Print rv0, " \\SD :0 Unknow device, 1 pass ,2 card change bit fail"
              
                 
                If TestResult = "PASS" Then
                     TestResult = ""
                    
                        CardResult = DO_WritePort(card, Channel_P1A, &H0)
                        Call MsecDelay(1#)
           
                    
                 
                 '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                 '
                 '  R/W test
                 '
                 '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                 
                
                'initial return value
                
                
                rv0 = 0
                rv1 = 0
                rv2 = 0
                rv3 = 0
                
                Label3.BackColor = RGB(255, 255, 255)
                Label4.BackColor = RGB(255, 255, 255)
                Label5.BackColor = RGB(255, 255, 255)
                Label6.BackColor = RGB(255, 255, 255)
                Label7.BackColor = RGB(255, 255, 255)
                Label8.BackColor = RGB(255, 255, 255)
                
                
                
                ClosePipe
                rv0 = CBWTest_New(0, 1, "vid_058f")
                Call LabelMenu(0, rv0, 1)
                ClosePipe
               
                
                 If rv0 = 1 Then
                                  CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                                 If LightOff <> 248 Then
                                   UsbSpeedTestResult = GPO_FAIL
                                    rv0 = 2
                                 End If
                 End If
                
                
                Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                 
                Print "LBA="; LBA
                
                
                'If rv0 = 1 And rv1 = 1 And rv2 = 1 And rv3 = 1 Then
                   ' TestResult = "PASS"
                'End If
                
                        
                        If rv0 = UNKNOW Then
                           UnknowDeviceFail = UnknowDeviceFail + 1
                           TestResult = "UNKNOW"
                        ElseIf rv0 = WRITE_FAIL Then
                            SDWriteFail = SDWriteFail + 1
                            TestResult = "SD_WF"
                        ElseIf rv0 = READ_FAIL Then
                            SDReadFail = SDReadFail + 1
                            TestResult = "SD_RF"
                        
                        ElseIf rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                        End If
                
               ' If rv2 * rv3 * rv1 = 0 Then
                '    MsgBox "rturn error"
                'End If
                  
                End If
                    
         Case "AU6333NC", "AU6331NC", "AU6333A1", "AU6333BL", "AU6333EL", "AU6376FLF20"
         
                Call AU6376FLF20TestSub
                'GPIO control setting
            
              
         Case "AU6333SD", "AU7630AF"
                  
                  
                If ChipName = "AU7630AN" Then
                   If PCI7248InitFinish = 0 Then
                      PCI7248Exist
                    End If
                    
                    CardResult = DO_WritePort(card, Channel_P1A, &H0)
                    Call MsecDelay(1#)
                     
                End If
                LBA = LBA + 1
                
                
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
                
                
                Print LBA
                
                ClosePipe
                rv0 = CBWTest_New_no_card(0, 1, "vid_058f")
                'Print "a1"
                Call LabelMenu(0, rv0, 1)
                ClosePipe
               
                
                
                
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv0 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                ElseIf rv0 = PASS Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                End If
                
                Print "Test Result"; TestResult
                       
               
              
               MSComm1.OutBufferCount = 0
               
               MSComm1.Output = TestResult   ' send out test result
            
                 
                Print rv0, " \\SD :0 Unknow device, 1 pass ,2 card change bit fail"
                
                 
                 
                If TestResult = "PASS" Then
                  
                     TestResult = ""
                      ChipName = ""
                        Do
                     
                              Call MsecDelay(0.1)
                              DoEvents
                              buf = MSComm1.Input
                              ChipName = ChipName & buf
                     
                        Loop Until InStr(1, ChipName, "PASS") <> 0
                 
                 
                 '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                 '
                 '  R/W test
                 '
                 '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                 
                
                'initial return value
                
                
                rv0 = 0
                rv1 = 0
                rv2 = 0
                rv3 = 0
                
                Label3.BackColor = RGB(255, 255, 255)
                Label4.BackColor = RGB(255, 255, 255)
                Label5.BackColor = RGB(255, 255, 255)
                Label6.BackColor = RGB(255, 255, 255)
                Label7.BackColor = RGB(255, 255, 255)
                Label8.BackColor = RGB(255, 255, 255)
                
                
                
                ClosePipe
                rv0 = CBWTest_New(0, 1, "vid_058f")
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                
                
                
                
                
                Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print " <<<<<<<<<<<<<<<<<<------------------>>>>>>>>>>>>>>>>>>>>>>"
                Print "LBA="; LBA
                
                
                'If rv0 = 1 And rv1 = 1 And rv2 = 1 And rv3 = 1 Then
                   ' TestResult = "PASS"
                'End If
                
                        
                        If rv0 = UNKNOW Then
                           UnknowDeviceFail = UnknowDeviceFail + 1
                           TestResult = "UNKNOW"
                        ElseIf rv0 = WRITE_FAIL Then
                            SDWriteFail = SDWriteFail + 1
                            TestResult = "SD_WF"
                        ElseIf rv0 = READ_FAIL Then
                            SDReadFail = SDReadFail + 1
                            TestResult = "SD_RF"
                        
                        ElseIf rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                        End If
                
               ' If rv2 * rv3 * rv1 = 0 Then
                '    MsgBox "rturn error"
                'End If
                End If
                
            Case "AU33SD_F"
                  
                     
                LBA = LBA + 1
                
                OldLBa = LBA
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
                
                
                Print LBA
                
                ClosePipe
                rv0 = CBWTest_New_no_card(0, 1, "vid_058f")
                'Print "a1"
                Call LabelMenu(0, rv0, 1)
                ClosePipe
               
                
                
                
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv0 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                ElseIf rv0 = PASS Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                End If
                
                Print "Test Result"; TestResult
                       
               
              
               MSComm1.OutBufferCount = 0
               
               MSComm1.Output = TestResult   ' send out test result
            
                 
                Print rv0, " \\SD :0 Unknow device, 1 pass ,2 card change bit fail"
                
                 
                 
                If TestResult = "PASS" Then
                  
                     TestResult = ""
                      ChipName = ""
                        Do
                     
                              Call MsecDelay(0.1)
                              DoEvents
                              buf = MSComm1.Input
                              ChipName = ChipName & buf
                     
                        Loop Until InStr(1, ChipName, "PASS") <> 0
                 
                 
                     '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                     '
                     '  R/W test
                     '
                     '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                     
                    
                    'initial return value
                    
                    
                    rv0 = 0
                    rv1 = 0
                    rv2 = 0
                    rv3 = 0
                    rv4 = 0
                    Label3.BackColor = RGB(255, 255, 255)
                    Label4.BackColor = RGB(255, 255, 255)
                    Label5.BackColor = RGB(255, 255, 255)
                    Label6.BackColor = RGB(255, 255, 255)
                    Label7.BackColor = RGB(255, 255, 255)
                    Label8.BackColor = RGB(255, 255, 255)
                    
                    
                    
                    ClosePipe
                    rv0 = CBWTest_New(0, 1, "vid_058f")
                    Call LabelMenu(0, rv0, 1)
                    ClosePipe
                    
                    rv4 = CBWTest_New_8_Sector(0, rv0, FailPosition, 20)
                    LBA = OldLBa
                    Call LabelMenu(0, rv4, rv0)
                    ClosePipe
                    
                    
                    
                    
                    Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                    Print " <<<<<<<<<<<<<<<<<<------------------>>>>>>>>>>>>>>>>>>>>>>"
                    Print rv4, " \\SD Stress:0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                    Print "LBA="; LBA
                    Print "FailPosition="; FailPosition
                    
                    
                    'If rv0 = 1 And rv1 = 1 And rv2 = 1 And rv3 = 1 Then
                       ' TestResult = "PASS"
                    'End If
                
                        
                    If rv0 = UNKNOW Then
                       UnknowDeviceFail = UnknowDeviceFail + 1
                       TestResult = "UNKNOW"
                    ElseIf rv0 = WRITE_FAIL Or rv4 = WRITE_FAIL Then
                        SDWriteFail = SDWriteFail + 1
                        TestResult = "SD_WF"
                    ElseIf rv0 = READ_FAIL Or rv4 = READ_FAIL Then
                        SDReadFail = SDReadFail + 1
                        TestResult = "SD_RF"
                    ElseIf rv1 = WRITE_FAIL Or rv5 = WRITE_FAIL Then
                        CFWriteFail = CFWriteFail + 1
                        TestResult = "CF_WF"
                    ElseIf rv1 = READ_FAIL Or rv5 = READ_FAIL Then
                        CFReadFail = CFReadFail + 1
                        TestResult = "CF_RF"
                    ElseIf rv2 = WRITE_FAIL Or rv6 = WRITE_FAIL Then
                        XDWriteFail = XDWriteFail + 1
                        TestResult = "XD_WF"
                    ElseIf rv2 = READ_FAIL Or rv6 = READ_FAIL Then
                        XDReadFail = XDReadFail + 1
                        TestResult = "XD_RF"
                    ElseIf rv3 = WRITE_FAIL Or rv7 = WRITE_FAIL Then
                        MSWriteFail = MSWriteFail + 1
                        TestResult = "MS_WF"
                    ElseIf rv3 = READ_FAIL Or rv7 = READ_FAIL Then
                        MSReadFail = MSReadFail + 1
                        TestResult = "MS_RF"
                    ElseIf rv4 * rv0 = PASS Then
                         TestResult = "PASS"
                    Else
                        TestResult = "Bin2"
                    End If
                
                End If
                
         Case "AU6336DFF21", "AU6336EFS10", "AU6366SD", "AU6331SD", "AU6332SD", "AU6332CF", "AU6371CF", "AU6332BSF0", "AU6371DFT10", "AU6331CSFT10", "AU6332BSF20", "AU6337BSF20", "AU6337CSF20", "AU6336AFF20", "AU6337CFF20", "AU6336DFF20", "AU6332GFF20", "AU6332FFF20", "AU6336ASF20", "AU6336EFF20", "AU6432BSF20"
         
              If ChipName = "AU6336ASF20" Then
              ChipName = "AU6337BSF20"
              End If
              
              If ChipName = "AU6336EFF20" Or ChipName = "AU6432BSF20" Then
        
              ChipName = "AU6336DFF20"
              End If
              
              If ChipName = "AU6336EFS10" Then
                PowerSet (4233)
                Call MsecDelay(0.6)
                ChipName = "AU6336DFF20"
              End If
             
               If ChipName = "AU6336DFF21" Then
                
                Call NotShareBusSingleSlotTestAU6336DFF21TestSub
                
                Else
                 Call NotShareBusSingleSlotTest
              End If
              
             
             
              
          Case "AU6336ZFF20"
               
               Call AU6336ZFF20TestSub
               
        Case "AU6336IFF01"
            Call AU6336IFF01TestSub
        
        Case "AU6336LFF01"
            Call AU6336LFF01TestSub
        
         Case "AU6336HFF21", "AU6336IFF21", "AU6336LFF21", "AU6336IFF20", "AU6336LFF20", "AU6336DLF20", "AU6336AAF20", "AU6336CAF20"
               
              If ChipName = "AU6336DLF20" Or ChipName = "AU6336AAF20" Then
                Call NotShareBusSingleSlotTestAU6336ExtRom
              ElseIf ChipName = "AU6336CAF20" Then
                Call NotShareBusSingleSlotTestAU6336CAF20
              
              ElseIf ChipName = "AU6336LFF21" Then
              
                 'Call NotShareBusSingleSlotTestAU6336LFF21TestSub
                 Call NotShareBusSingleSlotTestAU6336LFF21TestSub
              ElseIf ChipName = "AU6336IFF21" Then
                  Call NotShareBusSingleSlotTestAU6336IFF21TestSub
              ElseIf ChipName = "AU6336HFF21" Then
                  Call NotShareBusSingleSlotTestAU6336HFF21TestSub
              Else
                Call NotShareBusSingleSlotTestAU6336IF
              End If
   
      Case "AU6337BLF20", "AU6337GLF20", "AU6337ILF20", "AU6337BLF21", "AU6337BSF21", "AU6337ILF21", "AU6337GLF21", "AU6337CFF21", "AU6337CSF21"
         
            Call AU6337TestSub
            
              
   
   
   
        Case "AU6331MS"
            
                LBA = LBA + 1
                TestResult = ""
                ChipName = ""
                
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '
                '  R/W test
                '
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                
                'initial return value
                
                    rv0 = 0
                    rv1 = 0
                    rv2 = 0
                    rv3 = 0
                    
                    Label3.BackColor = RGB(255, 255, 255)
                    Label4.BackColor = RGB(255, 255, 255)
                    Label5.BackColor = RGB(255, 255, 255)
                    Label6.BackColor = RGB(255, 255, 255)
                    Label7.BackColor = RGB(255, 255, 255)
                    Label8.BackColor = RGB(255, 255, 255)
                
                ClosePipe
                    rv0 = CBWTest_New(0, 1, "vid_058f")
                    Call LabelMenu(3, rv0, 1)
                ClosePipe
                
                Print rv0, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print " <<<<<<<<<<<<<<<<<<------------------>>>>>>>>>>>>>>>>>>>>>>"
                Print "LBA="; LBA
                
                If rv0 = UNKNOW Then
                    UnknowDeviceFail = UnknowDeviceFail + 1
                    TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    MSWriteFail = MSWriteFail + 1
                    TestResult = "MS_WF"
                ElseIf rv0 = READ_FAIL Then
                    MSReadFail = MSReadFail + 1
                    TestResult = "MS_RF"
                ElseIf rv0 = PASS Then
                    TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                End If
                
            Case "AU6376BLF20", "AU6376BLF22", "AU6376FLF22"
            
                  Call AU6376TestSub
                  
                  
                  
                
                Case "AU6366CF", "AU6390", "AU6391BLF21"
            
                LBA = LBA + 1
                TestResult = ""
                
                If ChipName = "AU6391BLF21" Then
                Call MsecDelay(2.3)
                End If
                 
                
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '
                '  R/W test
                '
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                
                'initial return value
                
                rv0 = 0
                rv1 = 0
                rv2 = 0
                rv3 = 0
                rv4 = 0
                
                Label3.BackColor = RGB(255, 255, 255)
                Label4.BackColor = RGB(255, 255, 255)
                Label5.BackColor = RGB(255, 255, 255)
                Label6.BackColor = RGB(255, 255, 255)
                Label7.BackColor = RGB(255, 255, 255)
                Label8.BackColor = RGB(255, 255, 255)
                
                ClosePipe
                    rv0 = CBWTest_New(0, 1, "vid_058f")
                    Call LabelMenu(1, rv0, 1)
                ClosePipe
                
               
                Print rv0, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print " <<<<<<<<<<<<<<<<<<------------------>>>>>>>>>>>>>>>>>>>>>>"
                Print "LBA="; LBA
                
                If rv0 = UNKNOW Then
                    UnknowDeviceFail = UnknowDeviceFail + 1
                    TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    CFWriteFail = CFWriteFail + 1
                    TestResult = "CF_WF"
                ElseIf rv0 = READ_FAIL Then
                    CFReadFail = CFReadFail + 1
                    TestResult = "CF_RF"
                ElseIf rv0 = PASS Then
                    TestResult = "PASS"
                    
                   
                Else
                    TestResult = "Bin2"
                End If
                
            Case "AU6366XD", "AU6369XD"
  
                LBA = LBA + 1
                TestResult = ""
                ChipName = ""
                
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '
                '  R/W test
                '
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                
                'initial return value
                
                rv0 = 0
                rv1 = 0
                rv2 = 0
                rv3 = 0
                
                Label3.BackColor = RGB(255, 255, 255)
                Label4.BackColor = RGB(255, 255, 255)
                Label5.BackColor = RGB(255, 255, 255)
                Label6.BackColor = RGB(255, 255, 255)
                Label7.BackColor = RGB(255, 255, 255)
                Label8.BackColor = RGB(255, 255, 255)
                
                ClosePipe
                rv0 = CBWTest_New(0, 1, "vid_058f")
                Call LabelMenu(2, rv0, 1)
                ClosePipe
                
                Print rv0, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print " <<<<<<<<<<<<<<<<<<------------------>>>>>>>>>>>>>>>>>>>>>>"
                Print "LBA="; LBA
                
                If rv0 = UNKNOW Then
                UnknowDeviceFail = UnknowDeviceFail + 1
                TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                XDWriteFail = XDWriteFail + 1
                TestResult = "XD_WF"
                ElseIf rv0 = READ_FAIL Then
                XDReadFail = XDReadFail + 1
                TestResult = "XD_RF"
                
                ElseIf rv0 = PASS Then
                TestResult = "PASS"
                Else
                TestResult = "Bin2"
                End If
                
            Case "AU6366MS"
       
                LBA = LBA + 1
                TestResult = ""
                ChipName = ""
                
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '
                '  R/W test
                '
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                
                'initial return value
                
                rv0 = 0
                rv1 = 0
                rv2 = 0
                rv3 = 0
                
                Label3.BackColor = RGB(255, 255, 255)
                Label4.BackColor = RGB(255, 255, 255)
                Label5.BackColor = RGB(255, 255, 255)
                Label6.BackColor = RGB(255, 255, 255)
                Label7.BackColor = RGB(255, 255, 255)
                Label8.BackColor = RGB(255, 255, 255)
                
                ClosePipe
                rv0 = CBWTest_New(0, 1, "vid_058f")
                Call LabelMenu(3, rv0, 1)
                ClosePipe
                
                Print rv0, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print " <<<<<<<<<<<<<<<<<<------------------>>>>>>>>>>>>>>>>>>>>>>"
                Print "LBA="; LBA
                
                If rv0 = UNKNOW Then
                UnknowDeviceFail = UnknowDeviceFail + 1
                TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                MSWriteFail = MSWriteFail + 1
                TestResult = "MS_WF"
                ElseIf rv0 = READ_FAIL Then
                MSReadFail = MSReadFail + 1
                TestResult = "MS_RF"
                
                ElseIf rv0 = PASS Then
                TestResult = "PASS"
                Else
                TestResult = "Bin2"
                End If
                
            Case "AU6368_S"
                        
                LBA = LBA + 1
                ClosePipe
                rv3 = CBWTest_New(3, 1, "vid_058f")
                'rv3 = CBWTest(3, 1)
                Call LabelMenu(3, rv3, 1)
                ClosePipe
                               
                If rv3 = UNKNOW Then 'If rv3 = UNKNOW_DEVICE Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv3 = PASS Then 'Arch define 940410
                   TestResult = "PASS"
                ElseIf rv3 = WRITE_FAIL Then
                    MSWriteFail = MSWriteFail + 1
                    TestResult = "MS_WF"
                ElseIf rv3 = READ_FAIL Then
                    MSReadFail = MSReadFail + 1
                    TestResult = "MS_RF"
                Else
                    TestResult = "Bin2"
                End If
                        
                                
                Print rv3, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print "LBA="; LBA
               
          
               
               
                
             Case "AU6368A", "AU6368A1", "AU6370BL", "AU6375HL", "AU6375CL", "AU6376ELF20", "AU6376ILF20", "AU6375HLF21", "AU6370DLF20", "AU6378ALF20", "AU6376JLF20", "AU6377ALF21", "AU6377ALF24", "AU6377ALF25", "AU6370GLF20", "AU6375HLF22", "AU6375HLF23"
            
               
               Call MultiSlotTest
                    
               
            Case "AU6376ALO14", "AU6376ALOT1", "AU6376ALOT2", "AU6376ALOT3", "AU6376ALO13", "AU6376ALO12", "AU6376ALO11", "AU6376ALO10", "AU6376KLF20", "AU6376ALF20", "AU6376ALF21", "AU6376ALF22", "AU6376ELF22", "AU6376JLF22"
                   
               Call AU6376TestAllSub
                
                   
            Case "AU6378ALF23", "AU6378HLF23", "AU6378FLF22", "AU6378RLF23"
            
            If ChipName = "AU6378FLF22" Then
            MultiSlotTestAU6378FLTest
            End If
            
            If ChipName = "AU6378ALF23" Or ChipName = "AU6378HLF23" Then
            MultiSlotTestAU6378
            End If
            
             If ChipName = "AU6378RLF23" Then
            MultiSlotTestAU6378RLTestSub
            End If
            
            
             Case "AU6378ALF24", "AU6378HLF24", "AU6378FLF24", "AU6378RLF24", "AU6378ALF04", "AU6378ALS14"
             
            If ChipName = "AU6378ALF04" Then
                Call AU6378ALF04TestSub
            ElseIf ChipName = "AU6378ALS14" Then
                Call AU6378ALS14TestSub
            End If
             
            If ChipName = "AU6378FLF24" Then
                MultiSlotTestAU6378FLF24Test
            End If
            
            If ChipName = "AU6378ALF24" Or ChipName = "AU6378HLF24" Then
                MultiSlotTestAU6378ALF24
            End If
            
            If ChipName = "AU6378RLF24" Then
                MultiSlotTestAU6378RLF24TestSub
            End If
               
            Case "AU6363", "AU9368_1", "AU6368", "AU6375", "AU6375HLF20"
            
                 
                LBA = LBA + 1
                
            
               
                
                ClosePipe
                rv0 = CBWTest_New(0, 1, "vid_058f")
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                rv1 = CBWTest_New(1, rv0, "vid_058f")
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
                rv2 = CBWTest_New(2, rv1, "vid_058f")
                Call LabelMenu(2, rv2, rv1)
                ClosePipe
                rv3 = CBWTest_New(3, rv2, "vid_058f")
                Call LabelMenu(3, rv3, rv2)
                ClosePipe
                
                If ChipName = "AU6375HLF20" Then
                  
                ClosePipe
                rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                If rv0 = 1 Then
                   ClosePipe
                    rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                    Call LabelMenu(0, rv0, 1)
                    ClosePipe
                End If
               
                End If
                
                Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print rv3, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print "LBA="; LBA
                
                
                'If rv0 = 1 And rv1 = 1 And rv2 = 1 And rv3 = 1 Then
                   ' TestResult = "PASS"
                'End If
                
                
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv0 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                ElseIf rv1 = WRITE_FAIL Then
                    CFWriteFail = CFWriteFail + 1
                    TestResult = "CF_WF"
                ElseIf rv1 = READ_FAIL Then
                    CFReadFail = CFReadFail + 1
                    TestResult = "CF_RF"
                ElseIf rv2 = WRITE_FAIL Then
                    XDWriteFail = XDWriteFail + 1
                    TestResult = "XD_WF"
                ElseIf rv2 = READ_FAIL Then
                    XDReadFail = XDReadFail + 1
                    TestResult = "XD_RF"
                 ElseIf rv3 = WRITE_FAIL Then
                    MSWriteFail = MSWriteFail + 1
                    TestResult = "MS_WF"
                ElseIf rv3 = READ_FAIL Then
                    MSReadFail = MSReadFail + 1
                    TestResult = "MS_RF"
                ElseIf rv3 * rv2 * rv1 * rv0 = PASS Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
                
               ' If rv2 * rv3 * rv1 = 0 Then
                '    MsgBox "rturn error"
                'End If
                
                
                
                
            Case "AU6369S2", "AU6373S2"
            
            
                LBA = LBA + 1
                ClosePipe
                rv3 = CBWTest_New(3, 1, "vid_058f")
                'rv0 = CBWTest(0, 1)
                
                Call LabelMenu(3, rv3, 1)
                ClosePipe
                rv0 = CBWTest_New(0, rv3, "vid_058f")
                'rv3 = CBWTest(3, rv0)
                Call LabelMenu(0, rv0, rv3)
                'rv3 = CBWTest(3, 1)    'arch change 940415
                'Call LabelMenu(3, rv3, 1) 'arch change 940415
                ClosePipe
                'rv0 = CBWTest(0, rv3) 'arch change 940415
                'Call LabelMenu(0, rv0, rv3) 'arch change 940415
                             
                
                Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print rv3, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print "LBA="; LBA
                
                
                If rv3 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv0 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                ElseIf rv1 = WRITE_FAIL Then
                    CFWriteFail = CFWriteFail + 1
                    TestResult = "CF_WF"
                ElseIf rv1 = READ_FAIL Then
                    CFReadFail = CFReadFail + 1
                    TestResult = "CF_RF"
                ElseIf rv2 = WRITE_FAIL Then
                    XDWriteFail = XDWriteFail + 1
                    TestResult = "XD_WF"
                ElseIf rv2 = READ_FAIL Then
                    XDReadFail = XDReadFail + 1
                    TestResult = "XD_RF"
                ElseIf rv3 = WRITE_FAIL Then
                    MSWriteFail = MSWriteFail + 1
                    TestResult = "MS_WF"
                ElseIf rv3 = READ_FAIL Then
                    MSReadFail = MSReadFail + 1
                    TestResult = "MS_RF"
                ElseIf rv0 * rv3 = PASS Then
                     TestResult = "PASS"
                Else
                      TestResult = "Bin2"
                     
                End If
                
                
                If rv0 = 0 Then
                    MsgBox "return error"
                
                End If
                
                
                 Case "AU6334"
            
             
                LBA = LBA + 1
                ClosePipe
                
                rv1 = CBWTest_New(1, 1, "vid_058f")
                Call LabelMenu(1, rv1, 1) 'arch change 940415
  
                ClosePipe
                rv0 = CBWTest_New(0, rv1, "vid_058f")
                Call LabelMenu(0, rv0, rv1) 'arch change 940415
                ClosePipe
                rv2 = CBWTest_New(2, rv0, "vid_058f")
                Call LabelMenu(2, rv2, rv0) 'arch change 940415
                ClosePipe
                
                
                
                Print rv0, " \\MMC:0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print rv1, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print rv2, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print "LBA="; LBA
                
                
              If rv1 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv0 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                ElseIf rv1 = WRITE_FAIL Then
                    CFWriteFail = CFWriteFail + 1
                    TestResult = "CF_WF"
                ElseIf rv1 = READ_FAIL Then
                    CFReadFail = CFReadFail + 1
                    TestResult = "CF_RF"
                ElseIf rv2 = WRITE_FAIL Then
                    XDWriteFail = XDWriteFail + 1
                    TestResult = "XD_WF"
                ElseIf rv2 = READ_FAIL Then
                    XDReadFail = XDReadFail + 1
                    TestResult = "XD_RF"
                 ElseIf rv3 = WRITE_FAIL Then
                    MSWriteFail = MSWriteFail + 1
                    TestResult = "MS_WF"
                ElseIf rv3 = READ_FAIL Then
                    MSReadFail = MSReadFail + 1
                    TestResult = "MS_RF"
                ElseIf rv2 * rv0 * rv1 = PASS Then
                     TestResult = "PASS"
                Else
                     TestResult = "Bin2"
                     
                End If
                 
                 If rv0 * rv2 = 0 Then
                    MsgBox "return error"
                 End If
                
                
                
                Case "AU6334CLF20"
            
             
            
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                End If
                
                   result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
            
                    CardResult = DO_WritePort(card, Channel_P1B, &H0)
                    
               
                    CardResult = DO_WritePort(card, Channel_P1A, &HFF)  ' 1111 1110
                    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                    Call MsecDelay(0.3)
                    
                     CardResult = DO_WritePort(card, Channel_P1A, &H7F)  ' 1111 1110
                   
                Call MsecDelay(1)
               
               LBA = LBA + 1
                
                
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
                
                
                Print LBA
                
              
                ClosePipe
                 rv0 = CBWTest_New_no_card(0, 1, "vid_058f")
                'Print "a1"
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                rv1 = CBWTest_New_no_card(1, rv0, "vid_058f")
               '  Print "a2"
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
                
                rv2 = CBWTest_New_no_card(2, rv1, "vid_058f")
               '  Print "a3"
                Call LabelMenu(2, rv2, rv1)
          
                CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                If LightOff <> 255 Then
                     UsbSpeedTestResult = GPO_FAIL
                      rv0 = 2
                 End If
              Print rv0, " \\MMC:0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print rv1, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print rv2, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
          
                   
              If rv1 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv0 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                ElseIf rv1 = WRITE_FAIL Then
                    CFWriteFail = CFWriteFail + 1
                    TestResult = "CF_WF"
                ElseIf rv1 = READ_FAIL Then
                    CFReadFail = CFReadFail + 1
                    TestResult = "CF_RF"
                ElseIf rv2 = WRITE_FAIL Then
                    XDWriteFail = XDWriteFail + 1
                    TestResult = "XD_WF"
                ElseIf rv2 = READ_FAIL Then
                    XDReadFail = XDReadFail + 1
                    TestResult = "XD_RF"
                 ElseIf rv3 = WRITE_FAIL Then
                    MSWriteFail = MSWriteFail + 1
                    TestResult = "MS_WF"
                ElseIf rv3 = READ_FAIL Then
                    MSReadFail = MSReadFail + 1
                    TestResult = "MS_RF"
                ElseIf rv2 * rv0 * rv1 = PASS Then
                     TestResult = "PASS"
                Else
                     TestResult = "Bin2"
                     
                End If
             
            
              If TestResult = "PASS" Then
              
              
                    
                CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' 1111 1110
                   
                Call MsecDelay(0.3)
               
             
                LBA = LBA + 1
                ClosePipe
                
                rv1 = CBWTest_New(1, 1, "vid_058f")
                Call LabelMenu(1, rv1, 1) 'arch change 940415
  
                ClosePipe
                rv0 = CBWTest_New(0, rv1, "vid_058f")
                Call LabelMenu(0, rv0, rv1) 'arch change 940415
                ClosePipe
                rv2 = CBWTest_New(2, rv0, "vid_058f")
                Call LabelMenu(2, rv2, rv0) 'arch change 940415
                ClosePipe
                
                
                CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                If LightOff <> 247 Then
                     UsbSpeedTestResult = GPO_FAIL
                     rv0 = 2
                 End If
                
                Print rv0, " \\MMC:0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print rv1, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print rv2, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print "LBA="; LBA
                
                
              If rv1 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv0 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                ElseIf rv1 = WRITE_FAIL Then
                    CFWriteFail = CFWriteFail + 1
                    TestResult = "CF_WF"
                ElseIf rv1 = READ_FAIL Then
                    CFReadFail = CFReadFail + 1
                    TestResult = "CF_RF"
                ElseIf rv2 = WRITE_FAIL Then
                    XDWriteFail = XDWriteFail + 1
                    TestResult = "XD_WF"
                ElseIf rv2 = READ_FAIL Then
                    XDReadFail = XDReadFail + 1
                    TestResult = "XD_RF"
                 ElseIf rv3 = WRITE_FAIL Then
                    MSWriteFail = MSWriteFail + 1
                    TestResult = "MS_WF"
                ElseIf rv3 = READ_FAIL Then
                    MSReadFail = MSReadFail + 1
                    TestResult = "MS_RF"
                ElseIf rv2 * rv0 * rv1 = PASS Then
                     TestResult = "PASS"
                Else
                     TestResult = "Bin2"
                     
                End If
                 
               
                End If
                 
            Case "AU6369S3"
            
            
                LBA = LBA + 1
                ClosePipe
                
                   rv3 = CBWTest_New(3, 1, "vid_058f")
                    Call LabelMenu(3, rv3, 1) 'arch change 940415
  
                ClosePipe
                rv0 = CBWTest_New(0, rv3, "vid_058f")
                Call LabelMenu(0, rv0, rv3) 'arch change 940415
                ClosePipe
                rv2 = CBWTest_New(2, rv0, "vid_058f")
                Call LabelMenu(2, rv2, rv0) 'arch change 940415
                ClosePipe
                
                
                
                Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print rv3, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print "LBA="; LBA
                
                
              If rv3 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv0 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                ElseIf rv1 = WRITE_FAIL Then
                    CFWriteFail = CFWriteFail + 1
                    TestResult = "CF_WF"
                ElseIf rv1 = READ_FAIL Then
                    CFReadFail = CFReadFail + 1
                    TestResult = "CF_RF"
                ElseIf rv2 = WRITE_FAIL Then
                    XDWriteFail = XDWriteFail + 1
                    TestResult = "XD_WF"
                ElseIf rv2 = READ_FAIL Then
                    XDReadFail = XDReadFail + 1
                    TestResult = "XD_RF"
                 ElseIf rv3 = WRITE_FAIL Then
                    MSWriteFail = MSWriteFail + 1
                    TestResult = "MS_WF"
                ElseIf rv3 = READ_FAIL Then
                    MSReadFail = MSReadFail + 1
                    TestResult = "MS_RF"
                ElseIf rv2 * rv0 * rv3 = PASS Then
                     TestResult = "PASS"
                Else
                     TestResult = "Bin2"
                     
                End If
                 
                 If rv0 * rv2 = 0 Then
                    MsgBox "return error"
                 End If
                 
            Case "AU63693F"
                rv0 = 0
                rv1 = 0
                rv2 = 0
                rv3 = 0
                rv4 = 0
                rv5 = 0
                rv6 = 0
                rv7 = 0
            
                LBA = LBA + 1
                ClosePipe
                rv3 = CBWTest_New(3, 1, "vid_058f")
                Call LabelMenu(3, rv3, 1)
                ClosePipe
                
                rv0 = CBWTest_New(0, rv3, "vid_058f")
                Call LabelMenu(0, rv0, rv3)
                ClosePipe
                
                rv2 = CBWTest_New(2, rv0, "vid_058f")
                Call LabelMenu(2, rv2, rv0)
                ClosePipe
                
                LBA = OldLBa
                rv7 = CBWTest_New_8_Sector(3, rv2, FailPosition, 20)
                Call LabelMenu(3, rv7, rv2)
                ClosePipe
                
                LBA = OldLBa
                rv4 = CBWTest_New_8_Sector(0, rv7, FailPosition, 20)
                Call LabelMenu(0, rv4, rv7)
                ClosePipe
                
                LBA = OldLBa
                rv6 = CBWTest_New_8_Sector(2, rv4, FailPosition, 20)
                Call LabelMenu(2, rv6, rv4)
                ClosePipe
                
                Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print rv3, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                Print rv4, " \\SD stress :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print rv6, " \\MS stress :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print rv7, " \\MS stress :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print "LBA="; LBA
                
        
                If rv3 = UNKNOW Then
                    UnknowDeviceFail = UnknowDeviceFail + 1
                    TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Or rv4 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv0 = READ_FAIL Or rv4 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                ElseIf rv2 = WRITE_FAIL Or rv6 = WRITE_FAIL Then
                    XDWriteFail = XDWriteFail + 1
                    TestResult = "XD_WF"
                ElseIf rv2 = READ_FAIL Or rv6 = READ_FAIL Then
                    XDReadFail = XDReadFail + 1
                    TestResult = "XD_RF"
                ElseIf rv3 = WRITE_FAIL Or rv7 = WRITE_FAIL Then
                    MSWriteFail = MSWriteFail + 1
                    TestResult = "MS_WF"
                ElseIf rv3 = READ_FAIL Or rv7 = READ_FAIL Then
                    MSReadFail = MSReadFail + 1
                    TestResult = "MS_RF"
                ElseIf rv0 * rv2 * rv3 * rv4 * rv6 * rv7 = PASS Then
                    TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                End If
     
                 
            Case "AU63692F"
                
               
                LBA = LBA + 1
                 OldLBa = LBA
                ClosePipe
                      
                rv3 = CBWTest_New(3, 1, "vid_058f")
                
                
                Call LabelMenu(3, rv3, 1)
                ClosePipe
                rv0 = CBWTest_New(0, rv3, "vid_058f")
             
                Call LabelMenu(0, rv0, rv3)
              
                ClosePipe
               
                LBA = OldLBa
                rv7 = CBWTest_New_8_Sector(3, rv0, FailPosition, 20)
                LBA = OldLBa
                Call LabelMenu(3, rv7, rv0)
                ClosePipe
                rv4 = CBWTest_New_8_Sector(0, rv7, FailPosition, 20)
                 Call LabelMenu(0, rv4, rv7)
                ClosePipe
                
                Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print rv3, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print rv4, " \\SD stress :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print rv7, " \\MS stress :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print "LBA="; LBA
                Print "FailPosition="; FailPosition
                
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Or rv4 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv0 = READ_FAIL Or rv4 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                ElseIf rv1 = WRITE_FAIL Or rv5 = WRITE_FAIL Then
                    CFWriteFail = CFWriteFail + 1
                    TestResult = "CF_WF"
                ElseIf rv1 = READ_FAIL Or rv5 = READ_FAIL Then
                    CFReadFail = CFReadFail + 1
                    TestResult = "CF_RF"
                ElseIf rv2 = WRITE_FAIL Or rv6 = WRITE_FAIL Then
                    XDWriteFail = XDWriteFail + 1
                    TestResult = "XD_WF"
                ElseIf rv2 = READ_FAIL Or rv6 = READ_FAIL Then
                    XDReadFail = XDReadFail + 1
                    TestResult = "XD_RF"
                ElseIf rv3 = WRITE_FAIL Or rv7 = WRITE_FAIL Then
                    MSWriteFail = MSWriteFail + 1
                    TestResult = "MS_WF"
                ElseIf rv3 = READ_FAIL Or rv7 = READ_FAIL Then
                    MSReadFail = MSReadFail + 1
                    TestResult = "MS_RF"
                ElseIf rv7 * rv4 * rv3 * rv0 = PASS Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                     
                End If
                     
                 
            Case "AU698XELF21", "AU6985HLS10", "AU698XHLF20", "AU698XILF20", "AU698XHLF21", "AU698XILF21"
         
               Call AU698XTestSub
               
                
            
            Case "AU6330", "AU6385", "AU6386", "AU6386_D", "AU6385_1", "AU6385_2", "AU9386", "AU6388", "AU6389", "AU6980", "AU6980AN"
            
            '//*LabelMenu(SlotNo As Byte, TestResult As Byte, PreSlotStatus As Byte)
                
              Call AU6980TestSub
              
           Case "AU6980OCP10"
              Call AU6980OCP10TestSub
              
           Case "AU6980HLS10"
                Call AU6980HLSTestSub
              
            Case "AU6986HLF21", "AU6986ALF21"
            
                If Left(ChipName, 8) = "AU6986HL" Then
              
                 Call AU6986TestSub
               Else
                 Call AU6986ALTestSub
               End If
               
            
            Case "AU6388F", "AU6386DF", "AU6389F", "AU6980F"
                
                '//*LabelMenu(SlotNo As Byte, TestResult As Byte, PreSlotStatus As Byte)
                
                LBA = LBA + 1
                OldLBa = LBA
                ClosePipe
                rv0 = CBWTest_New(0, 1, "vid_058f")
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                
                LBA = OldLBa
                ClosePipe
                rv4 = CBWTest_New_8_Sector(0, rv0, FailPosition, 20)
                Call LabelMenu(0, rv4, rv0)
                ClosePipe
                
                
                
                Print rv0, " Flash \\ : 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print rv4, " Flash stress\\ : 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print "LBA="; LBA
                Print "FailPosition="; FailPosition
                
                If rv0 = UNKNOW Then
                    UnknowDeviceFail = UnknowDeviceFail + 1
                    TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Or rv4 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv0 = READ_FAIL Or rv4 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                ElseIf rv4 * rv0 = PASS Then
                    TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                End If
                
            Case "AU9368"
            
                LBA = LBA + 1
                OldLBa = LBA
                ClosePipe
                rv0 = CBWTest_New(0, 1, "vid_058f")
                Call LabelMenu(0, rv0, 1)
                '//*LabelMenu(SlotNo As Byte, TestResult As Byte, PreSlotStatus As Byte)
                ClosePipe
                rv1 = CBWTest_New(1, rv0, "vid_058f")
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
                rv2 = CBWTest_New(2, rv1, "vid_058f")
                Call LabelMenu(2, rv2, rv1)
                ClosePipe
                rv3 = CBWTest_New(3, rv2, "vid_058f")
                Call LabelMenu(3, rv3, rv2)
                ClosePipe
                rv4 = CBWTest_New_8_Sector(2, rv3, FailPosition, 64)
                LBA = OldLBa
                Call LabelMenu(2, rv4, rv3)
                ClosePipe
                rv5 = CBWTest_New_8_Sector(3, rv4, FailPosition, 64)
                LBA = OldLBa
                Call LabelMenu(3, rv5, rv4)
                ClosePipe
                Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print rv3, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print rv4, " \\XD stress :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print rv5, " \\MS stress :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Print "LBA="; LBA
                Print "FailPosition="; FailPosition
                
                
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv0 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                ElseIf rv1 = WRITE_FAIL Then
                    CFWriteFail = CFWriteFail + 1
                    TestResult = "CF_WF"
                ElseIf rv1 = READ_FAIL Then
                    CFReadFail = CFReadFail + 1
                    TestResult = "CF_RF"
                ElseIf rv2 = WRITE_FAIL Or rv4 = WRITE_FAIL Then
                    XDWriteFail = XDWriteFail + 1
                    TestResult = "XD_WF"
                ElseIf rv2 = READ_FAIL Or rv4 = READ_FAIL Then
                    XDReadFail = XDReadFail + 1
                    TestResult = "XD_RF"
                ElseIf rv3 = WRITE_FAIL Or rv5 = WRITE_FAIL Then
                    MSWriteFail = MSWriteFail + 1
                    TestResult = "MS_WF"
                ElseIf rv3 = READ_FAIL Or rv5 = READ_FAIL Then
                    MSReadFail = MSReadFail + 1
                    TestResult = "MS_RF"
                ElseIf rv5 * rv4 * rv3 * rv2 * rv2 * rv0 = PASS Then
                     TestResult = "PASS"
                End If
                
             '    If rv5 * rv4 * rv3 * rv2 * rv1 = 0 Then
             '       MsgBox "return error"
            '    End If
                
                
                 
           
                
             Case "AU9520"
                PreviousStatus = 0
                TestResultadd = "FAIL"
                TestResultSmartCard = "FAIL"
                ArchTest = False
                
                Cls
                ' txtmsg.Text = ""
                
                strReaderName = VenderString & " " & ProductString & " 0"   'slot 0
                udtReaderStates(0).szReader = strReaderName & vbNullChar
                txtmsg.Text = udtReaderStates(0).szReader
                
                Print "Start to test slot0 !"
                Call StartTest
                
                If ArchTest = True Then
                    Print "slot0 test ok!"
                    
                    strReaderName = VenderString & " " & ProductString & " 1"   'slot 1
                    udtReaderStates(0).szReader = strReaderName & vbNullChar
                    txtmsg.Text = udtReaderStates(0).szReader
                    
                    Print "Start to test slot1 !"
                    Call StartTest
                    Print "slot1 test ok!"
                Else
                    Print "slot1 not test !"
                End If
                
                'If TestResultSmartCard = "PASS" Then
               
                '    TestResult = "PASS"
                'Else
                  '  TestResult = "FAIL"
                'End If
                
                If ArchTest = True Then      'Arch define 940410
                    If TestResultSmartCard = "PASS" Then
                        TestResult = "PASS"
                    Else
                        TestResult = "Bin3"
                    End If
                Else
                    TestResult = "Bin2"
                End If
                
                
           ' Case "AU9720"
            
               ' Dim trans1 As Integer
               ' Dim trans2 As Integer
                
                  '  Text1.Text = ""
                   ' Text2.Text = ""
                'Cls
                   ' MSComm2.PortOpen = False
                   ' MSComm3.PortOpen = False
               ' Call MsecDelay(0.1)
                   ' MSComm2.PortOpen = True
                   ' MSComm3.PortOpen = True
                'Call MsecDelay(0.1)
                
                    'MSComm2.Output = "ArchTestFirst"
                
               ' Call MsecDelay(0.03)
                
                  '  Text1.Text = MSComm3.Input
                
               ' If Text1.Text = "ArchTestFirst" Then
                  '  trans1 = 1
                'Else
                  '  trans1 = 0
               ' End If
                
                   ' MSComm3.Output = "ArchTestSecond"
                
               ' Call MsecDelay(0.03)
                
                   ' Text2.Text = MSComm2.Input
                
               ' If Text2.Text = "ArchTestSecond" Then
                  '  trans2 = 1
               ' Else
                '    trans2 = 0
               ' End If
                
               ' If trans1 = 1 And trans2 = 1 Then
                    ' AU9720Test = 1
                 '   TestResult = "PASS"
               ' End If
            Case "AU6258Reg10", "AU6258Reg11"
            
                AU6258RegTestSub
            
            Case "AU6257URF20"
                
                AU6257URF20TestSub
            
            Case "AU6259KFS10", "AU6259KFS11", "AU6259KFS20"
                
                AU6259_SortingSub

            Case "AU6259BFS10"
                
                AU6259BFS10TestSub
            
            Case Else
                
                Label1.Caption = "Comm Err!!"
                
                
                Print "Comm Err!! ChipName="; ChipName
                MsgBox "Comm err!1"
                
            
        End Select
        
            
                      
      Print "Previous TestTime:"; Timer - OldTime
      'MsgBox Timer - oldtime
            
      'If TestResult = "PASS" Then
         
        ' Label2 = "PASS!"
         
        ' Label8.BackColor = RGB(0, 255, 0)
         'goodchip = goodchip + 1
      'Else
    
         'Label2 = "fail!"
         
         'Label8.BackColor = RGB(255, 0, 0)
         'failchip = failchip + 1
      'End If
        If TestResult = "PASS" Then  'arch change 940411
        
            Label2 = "Bin1, PASS!"
            Label8.BackColor = RGB(0, 255, 0)
            goodchip = goodchip + 1
            ContiUnknowFailCounter = 0
            
        ElseIf TestResult = "Bin2" Or TestResult = "UNKNOW" Then
        
            failchipBin2 = failchipBin2 + 1
            
            '2012/07/17 Reset Root-Hub if over 5 pcs
            ContiUnknowFailCounter = ContiUnknowFailCounter + 1
            
            If ContiUnknowFailCounter >= 5 Then
                ResetHubReturn = Shell(ResetHubString, vbNormalFocus)
                WaitProcQuit (ResetHubReturn)
                ContiUnknowFailCounter = 0
                MsecDelay (2#)
            End If
            
            Label2 = "Bin2 ,Unknow device "
            Label8.BackColor = RGB(255, 0, 0)
            failchip = failchip + 1
            TestResult = "Bin2"
        ElseIf TestResult = "Bin3" Or TestResult = "SD_WF" Or TestResult = "SD_RF" Or TestResult = "CF_WF" Or TestResult = "CF_RF" Then
        
            failchipBin3 = failchipBin3 + 1
            Label2 = "Bin3,SDfail!,CF fail, speed error!"
            Label8.BackColor = RGB(255, 0, 0)
            failchip = failchip + 1
            TestResult = "Bin3"
            
            '2013/04/02 Reset Root-Hub if over 5 pcs for FT6
            ContiUnknowFailCounter = ContiUnknowFailCounter + 1
            
            If ContiUnknowFailCounter >= 5 Then
                ResetHubReturn = Shell(ResetHubString, vbNormalFocus)
                WaitProcQuit (ResetHubReturn)
                ContiUnknowFailCounter = 0
                MsecDelay (2#)
            End If
        ElseIf TestResult = "Bin4" Or TestResult = "XD_WF" Or TestResult = "XD_RF" Then
        
            failchipBin4 = failchipBin4 + 1
            Label2 = "Bin4, XD fail!, AU6254 down stream port unknow"
            Label8.BackColor = RGB(255, 0, 0)
            failchip = failchip + 1
            TestResult = "Bin4"
            ContiUnknowFailCounter = 0
        ElseIf TestResult = "Bin5" Or TestResult = "MS_WF" Or TestResult = "MS_RF" Then
        
            failchipBin5 = failchipBin5 + 1
            Label2 = "Bin5 ,MS fail! AU6254 R/W fail, GPO fail       "
            Label8.BackColor = RGB(255, 0, 0)
            failchip = failchip + 1
            TestResult = "Bin5"
            ContiUnknowFailCounter = 0
         
            
        Else
        
            Label2 = "bin fail!, "
            Label8.BackColor = RGB(255, 0, 0)
            failchip = failchip + 1
            
        End If
  
    MSComm1.Output = TestResult
    Call MsecDelay(0.1) 'arch add
    Print "TestResulr :"; TestResult
     
    Text3.Text = goodchip
    Text4.Text = failchip
    Text5.Text = failchipBin2 'arch add 940411
    Text6.Text = failchipBin3 'arch add 940411
    Text7.Text = failchipBin4 'arch add 940411
    Text8.Text = failchipBin5 'arch add 940411
    
    '\\\\\\\\\\\\\\\\\
    ' Software Bin Counter
    '\\\\\\\\\\\\\\\\\
    
     Text9.Text = UnknowDeviceFail
    Text13.Text = SDWriteFail
    Text14.Text = SDReadFail
    Text15.Text = CFWriteFail
    Text16.Text = CFReadFail
    
    Text17.Text = XDWriteFail
    Text18.Text = XDReadFail
    
    Text21.Text = MSWriteFail
    Text22.Text = MSReadFail
  
    
    MSComm1.InBufferCount = 0
    MSComm1.InputLen = 0
    
    
 '   If ChipName = "AU6254XLS40" Then
  '     Call MsecDelay(3)
   '    Shell "shutdown -s -f -t 0"
    'End If
'========================

   ' Do
       'MSComm1.Output = "Ready"
        
        'Print TesterReadyCounter
       ' DoEvents
        'Call MsecDelay(0.1) ' comm time
      ' Endmsg = MSComm1.Input
        fnScsi2usb2K_KillEXE ' clear removable device message box
        
   ' Loop Until InStr(1, Endmsg, "testend") <> 0
       
    ' Endmsg = ""
    
    If bNeedsReStart Then
        CheckAvailMemSize (110)
    End If
    
     
Loop While (AllenTest = 0) 'And (Not AtheistDebug)


End Sub




'Private Sub cmdAT45D041_Click()
'fmAT45D041_1.Show vbModal
'End Sub


'Private Sub cmdClose_Click()
'bReadyToClose = True
'Unload fmMain
'End Sub











'Private Sub cmdStartTest_Click()
  '  cmdStartTest.Enabled = False
  '  cmdStopTest.Enabled = True
  '  cmdClose.Enabled = False

  '  Call StartTest
'End Sub

'Private Sub cmdStopTest_Click()
'bStop = True
'cmdStartTest.Enabled = True
'cmdStopTest.Enabled = False
'cmdClose.Enabled = True

'End Sub




Public Sub Form_Unload(Cancel As Integer)

mmcAudio.Command = "Close"
    Dim lngResult As Long
   
    If (bReadyToClose = False) Then
        Cancel = True
        Exit Sub
    End If
  '  If Left(ChipName, 8) = "AU6476BL" Then
   ' fnFreeDeviceHandle (DeviceHandle)
  '  End If
    'Scarddisconnect will cold reset the card and then power down the card.
    'if the EEPROM card presents, the function will take several seconds to cold reset it 3 times(because of time out).
    'for save time to close the program, this function can be deleted from the code, and the pc/sc subsystem will
    ' automatically do the function after the program is terminated.
    lngResult = SCardDisconnect(lngCard, SCARD_LEAVE_CARD) 'disconnect the connection with the reader
    lngResult = SCardReleaseContext(lngContext) ' release the connection with the resource manager
End Sub

'Public Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
 '   Select Case Chr(KeyCode)
  '      Case "R", "r"
'
 '           If cmdStartTest.Enabled = True Then Call cmdStartTest_Click

  '      Case "S", "s"
   '         If cmdStopTest.Enabled = True Then Call cmdStopTest_Click

   '     Case "X", "x"
    '        If cmdClose.Enabled = True Then Call cmdClose_Click

   ' End Select
'End Sub
'Public Sub Form_Keydown(KeyCode As Integer, Shift As Integer)
 '   Select Case Chr(KeyCode)
  '      Case "R", "r"
            
   '         If cmdStartTest.Enabled = True Then cmdStartTest.SetFocus
        
    '    Case "S", "s"
     '       If cmdStopTest.Enabled = True Then cmdStopTest.SetFocus
            
      '  Case "X", "x"
       '     If cmdClose.Enabled = True Then cmdClose.SetFocus
            
   ' End Select

'End Sub






Public Function AU6610Test() As Integer
On Error Resume Next
AU6610Test = Test

End Function






















Private Sub mmcAudio_Done(NotifyCode As Integer)
    mmcAudio.Command = "Close"
    'lblAudio.Caption = ""
    cmdPlay.Caption = "Play"
End Sub

Private Sub MP3_Rec_Click()

Print "begin record Mp3"
Dim i As Integer

                ChipName = "AU3130BL" '======================= need change
                'ChipName = "AU3150JLF20" '(1)======================= 2007.06.14 CHEYENNE CHANG
                RecordMode = 9
                RecordMode = 13
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                End If
    
    
                   result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
                  CardResult = DO_WritePort(card, Channel_P1B, &HFF)
                Call MsecDelay(10)
                  CardResult = DO_WritePort(card, Channel_P1B, &H0)
                Call MsecDelay(5)
                
                CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                             
                 If CardResult <> 0 Then
                       MsgBox "Power off fail"
                       End
                  End If
                            ' mp3==================================================
                  Call MsecDelay(0.5)
                             
                  CardResult = DO_WritePort(card, Channel_P1A, &H9F)   ' MP3 Power Off
                               
                  Call MsecDelay(4)    'power on time and DAC initial time
                             
                             
                     CardResult = DO_WritePort(card, Channel_P1A, &H8F)   ' MP3 Power Off
                            ' Call MsecDelay(2) '-- for CL
                         
                           '  Call MsecDelay(0.5) '-- for BL
                             
                          '   Call MsecDelay(0.01)
                          
                              CardResult = DO_WritePort(card, Channel_P1A, &H8D)
                            
                           '  Call MsecDelay(10)    '-- for AU3150JL play
                    
                     If CardResult <> 0 Then
                       MsgBox "Power on fail"
                       End
                    End If
                             
                             
               Call SetVol
                 mmcAudio.FileName = "D:\Documents and Settings\Administrator\桌面\New_host_PCI7248_2_960821-10\MP3 Pattern\test music4.mp3"
       mmcAudio.Command = "Open"
        mmcAudio.Command = "Play"
      
     '    cmdPlay.Caption = "Stop"
            ' Call MsecDelay(3)
                   i = AU3130Test
                 
                 
                '   Open App.Path & "\MP3_AU62541.txt" For Output As #4  '==========2007.06.14 CHEYENNE CHANG

                 '    For i = 0 To 100
                 '       Print #4, gnBuffer(i)
                '     Next i
                  

                 '   Close #4
                    
                  Call MsecDelay(20)
                '  Print "Test 20s"
                   i = AU3130Test
              
                 
                 '   Open App.Path & "\MP3.txt" For Output As #4
                    Open App.Path & "\MP3_AU62542.txt" For Output As #4  '==========2007.06.14 CHEYENNE CHANG

                     For i = 0 To 100
                        Print #4, gnBuffer(i)
                     Next i
                  

                    Close #4
                    
                  
                    MsgBox "Recode MP3 ok"
End Sub

Private Sub Rec_MP31_Click()
Print "begin record Mp31"
Dim i As Integer

                RecordMode = 1
 ChipName = "AU3130B"
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                End If
              
               
             
                CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                             
                 If CardResult <> 0 Then
                       MsgBox "Power off fail"
                       End
                  End If
                            ' mp3==================================================
                  Call MsecDelay(0.5)
                             
                  CardResult = DO_WritePort(card, Channel_P1A, &H9F)   ' MP3 Power Off
                               
                  Call MsecDelay(4)    'power on time and DAC initial time
                             
                             
                      CardResult = DO_WritePort(card, Channel_P1A, &H8F)   ' MP3 Power Off
                          '  Call MsecDelay(2)
                           '  Call MsecDelay(0.5)
                              CardResult = DO_WritePort(card, Channel_P1A, &H8D)
                            
                           '  Call MsecDelay(3)    'play
                    
                     If CardResult <> 0 Then
                       MsgBox "Power on fail"
                       End
                    End If
                             
               
                   i = AU3130Test
              
                 
                    Open App.Path & "\MP312.txt" For Output As #4

                     For i = 0 To 100
                        Print #4, gnBuffer(i)
                     Next i
                  

                    Close #4

                    MsgBox "Recode MP31 ok"
End Sub

Private Sub Rec_MP32_Click()
Print "begin record Mp32"
Dim i As Integer
 ChipName = "AU3130B"
                RecordMode = 1

                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                End If
              
               
             
                CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                             
                 If CardResult <> 0 Then
                       MsgBox "Power off fail"
                       End
                  End If
                            ' mp3==================================================
                  Call MsecDelay(0.5)
                             
                  CardResult = DO_WritePort(card, Channel_P1A, &H9F)   ' MP3 Power Off
                               
                  Call MsecDelay(4)    'power on time and DAC initial time
                             
                             
                      CardResult = DO_WritePort(card, Channel_P1A, &H8F)   ' MP3 Power Off
                          '  Call MsecDelay(2)
                             Call MsecDelay(0.5)
                              CardResult = DO_WritePort(card, Channel_P1A, &H8D)
                            
                           '  Call MsecDelay(3)    'play
                    
                     If CardResult <> 0 Then
                       MsgBox "Power on fail"
                       End
                    End If
                             
               
                   i = AU3130Test
              
                 
                    Open App.Path & "\MP32.txt" For Output As #4

                     For i = 0 To 100
                        Print #4, gnBuffer(i)
                     Next i
                  

                    Close #4

                    MsgBox "Recode MP32 ok"
End Sub

Private Sub Rec_Mp33_Click()
Print "begin record Mp33"
Dim i As Integer
 ChipName = "AU3130B"
                RecordMode = 1

                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                End If
              
               
             
                CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                             
                 If CardResult <> 0 Then
                       MsgBox "Power off fail"
                       End
                  End If
                            ' mp3==================================================
                  Call MsecDelay(0.5)
                             
                  CardResult = DO_WritePort(card, Channel_P1A, &H9F)   ' MP3 Power Off
                               
                  Call MsecDelay(4)    'power on time and DAC initial time
                             
                             
                      CardResult = DO_WritePort(card, Channel_P1A, &H8F)   ' MP3 Power Off
                           ' Call MsecDelay(2)
                             Call MsecDelay(0.5)
                              CardResult = DO_WritePort(card, Channel_P1A, &H8D)
                            
                           '  Call MsecDelay(3)    'play
                    
                     If CardResult <> 0 Then
                       MsgBox "Power on fail"
                       End
                    End If
                             
               
                   i = AU3130Test
              
                 
                    Open App.Path & "\MP33.txt" For Output As #4

                     For i = 0 To 100
                        Print #4, gnBuffer(i)
                     Next i
                  

                    Close #4

                    MsgBox "Recode MP33 ok"
End Sub

Private Sub Rec_Mp34_Click()
Print "begin record Mp34"
Dim i As Integer
               ChipName = "AU3130B"
                RecordMode = 1

                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                End If
              
               
             
                CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                             
                 If CardResult <> 0 Then
                       MsgBox "Power off fail"
                       End
                  End If
                            ' mp3==================================================
                  Call MsecDelay(0.5)
                             
                  CardResult = DO_WritePort(card, Channel_P1A, &H9F)   ' MP3 Power Off
                               
                  Call MsecDelay(4)    'power on time and DAC initial time
                             
                             
                      CardResult = DO_WritePort(card, Channel_P1A, &H8F)   ' MP3 Power Off
                          '  Call MsecDelay(2)
                            Call MsecDelay(0.5)
                              CardResult = DO_WritePort(card, Channel_P1A, &H8D)
                            
                           '  Call MsecDelay(3)    'play
                    
                     If CardResult <> 0 Then
                       MsgBox "Power on fail"
                       End
                    End If
                             
               
                   i = AU3130Test
              
                 
                    Open App.Path & "\MP34.txt" For Output As #4

                     For i = 0 To 100
                        Print #4, gnBuffer(i)
                     Next i
                  

                    Close #4

                    MsgBox "Recode MP34 ok"
End Sub

Private Sub Timer1_Timer()
Dim hwnd As Long, hdc As Long, X As Long, Y As Long
Dim lpPoint As POINTAPI, lpRect As RECT
Dim BColor As Long
'GetCursorPos lpPoint '取得滑鼠座標
lpPoint.X = 609
lpPoint.Y = 231

hwnd = WindowFromPoint(lpPoint.X, lpPoint.Y) '由回屬座標取得視窗的hWnd
GetWindowRect hwnd, lpRect '取得視窗的範圍
X = lpPoint.X - lpRect.Left '這是相對於視窗的x座標
Y = lpPoint.Y - lpRect.Top '這是相對於視窗的y座標

'取得視窗的hDC
hdc = GetWindowDC(hwnd)

'取得顏色並顯示
 BColor = GetPixel(hdc, X, Y)
 Picture1.BackColor = BColor
 
'Print Hex(BColor)
ReleaseDC hwnd, hdc
'Print BColor, Hex(BColor)
'can not use hex format
If BColor = 65535 Then
   CMediaTestResult = 1
End If
End Sub

Private Sub Timer2_Timer()
    If startProgress Then
        ProgressBar1.Value = ProgressBar1.Value + 2.5
        
        If ProgressBar1.Value >= ProgressBar1.Max Then
            Timer2.Enabled = False
        End If
    End If
End Sub
