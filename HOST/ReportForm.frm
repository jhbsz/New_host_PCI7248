VERSION 5.00
Begin VB.Form ReportForm 
   Caption         =   "報表設定"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8775
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6810
   ScaleWidth      =   8775
   StartUpPosition =   3  '系統預設值
   Begin VB.ComboBox Combo1 
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
      TabIndex        =   12
      Top             =   4920
      Width           =   3255
   End
   Begin VB.CommandButton MakeSure 
      Caption         =   "確定"
      Height          =   615
      Left            =   3600
      TabIndex        =   10
      Top             =   5880
      Width           =   1815
   End
   Begin VB.TextBox OperatorNameText 
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
      Left            =   4440
      TabIndex        =   9
      Top             =   3960
      Width           =   3135
   End
   Begin VB.TextBox HandlerIDText 
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
      Left            =   4440
      TabIndex        =   8
      Top             =   3120
      Width           =   3135
   End
   Begin VB.TextBox LotIDText 
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
      Left            =   4440
      TabIndex        =   7
      Top             =   2280
      Width           =   3135
   End
   Begin VB.TextBox RunCardNOText 
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
      Left            =   4440
      TabIndex        =   6
      Top             =   1320
      Width           =   3135
   End
   Begin VB.TextBox DeviceIDText 
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
      Left            =   4440
      TabIndex        =   5
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label Label6 
      Caption         =   "Process :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   11
      Top             =   4920
      Width           =   2415
   End
   Begin VB.Label Label5 
      Caption         =   "Handle ID:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   3240
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "Operator Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   4080
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Cust Lot ID:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Run Card NO:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Device ID:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   600
      Width           =   2295
   End
End
Attribute VB_Name = "ReportForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Sub InsertDB()
Dim oDB As ADOX.Catalog
Dim sDBPAth As String
Dim sConStr As String
Dim oCn As ADODB.Connection
Dim oCM As ADODB.Command
Dim RS As ADODB.Recordset

'-----------------------------
' set Path and connection string
'---------------------------
sDBPAth = "D:\SLT Summary\Summary.mdb"
'Debug.Print "1"; Dir(sDBPAth, vbNormal + vbDirectory)
If Dir(sDBPAth, vbNormal + vbDirectory) = " " Then
    MsgBox "MDB no EXIST"
    Exit Sub
End If

EndAt = "NA"

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

Set RS = oCn.Execute("Summary")
If RS.Fields("DeviceID").DefinedSize <= 50 Then
    Set RS = Nothing
    oCn.Execute "ALTER TABLE Summary alter column DeviceID char(50)"
End If

Set oCM = New ADODB.Command
oCM.ActiveConnection = oCn


 Dim cmstr As String

'cmstr = "INSERT INTO Summary VALUES(" & _

'oCM.CommandText = "INSERT INTO Summary VALUES(" & _

'cmstr = "INSERT INTO Summary VALUES(" & _

oCM.CommandText = "INSERT INTO Summary VALUES(" & _
"'" & NameofPC & " '," & _
"'" & ProgramName & "' ," & _
"'" & ProgramRevisionCode & "' ," & _
"'" & DeviceID & "' ," & _
"'" & RunCardNO & "' ," & _
"'" & ProcessID & "' ," & _
"'" & LotID & "' ," & _
"'" & StartAt & "' ," & _
"'" & EndAt & "' ," & _
"'" & HandlerID & "' ," & _
"'" & OperatorName & "' ," & _
Bin1Site1 & " ," & _
Bin1Site2 & " ," & _
Bin2Site1 & " ," & _
Bin2Site2 & " ," & _
Bin3Site1 & " ," & _
Bin3Site2 & " ," & _
Bin4Site1 & " ," & _
Bin4Site2 & " ," & _
Bin5Site1 & " ," & _
Bin5Site2 & _
")"

Debug.Print cmstr
oCM.Execute

 
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
Sub CreateDB()
Dim oDB As ADOX.Catalog
Dim sDBPAth As String
Dim sConStr As String
Dim oCn As ADODB.Connection
Dim oCM As ADODB.Command

'-----------------------------
' set Path and connection string
'---------------------------
sDBPAth = "D:\SLT Summary"
 Debug.Print "1"; Dir(sDBPAth, vbNormal + vbDirectory)
 If Dir(sDBPAth, vbNormal + vbDirectory) = "" Then
     MkDir sDBPAth
 End If


If Dir(sDBPAth & "\SLT.mdb", vbNormal + vbDirectory) <> "" Then
   Exit Sub
End If
 

'sConStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & "D:\SLT Summary" & "\SLT.mdb;Persist   Security   Info=False;Jet   OLEDB:Database   Password=058f"
sConStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & "D:\SLT Summary" & "\SLT.mdb"
 
' ------------------------
' Create New ADOX Object
' ------------------------
Set oDB = New ADOX.Catalog
oDB.Create sConStr

Set oCn = New ADODB.Connection
oCn.ConnectionString = sConStr
oCn.Open

Set oCM = New ADODB.Command
oCM.ActiveConnection = oCn

oCM.CommandText = "Create Table Summary (" & _
"[NameofPC] Text(20), " & _
"[ProgramName] Text(20), " & _
"[ProgramRevisionCode] Text(20), " & _
"[DeviceID] Text(50), " & _
"[RunCardNO] Text(20), " & _
"[ProcessID] Text(20), " & _
"[LotID] Text(20), " & _
"[StartAt] Text(20), " & _
"[EndAt] Text(20), " & _
"[HandlerID] Text(20), " & _
"[OperatorName] Text(20), " & _
"[Bin1Site1] Int," & _
"[Bin1Site2] Int," & _
"[Bin2Site1] Int," & _
"[Bin2Site2] Int," & _
"[Bin3Site1] Int," & _
"[Bin3Site2] Int," & _
"[Bin4Site1] Int," & _
"[Bin4Site2] Int," & _
"[Bin5Site1] Int," & _
"[Bin5Site2] Int" & _
")"


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
If err <> 0 Then
err.Clear
Resume Next
End If
End Sub


Private Sub DeviceIDText_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyReturn) Then
        RunCardNOText.SetFocus
    End If
End Sub

Private Sub Form_Load()

If AllenDebug = 1 Then
  DeviceIDText.Text = "AU6471b51-JFL-GR"
  RunCardNOText.Text = "2008CLOP0040"
  LotIDText.Text = "A697940.1"
  HandlerIDText.Text = "NS-1"
  OperatorNameText.Text = "149"
End If



Call CreateDB



Dim len5 As Long, aa As Long
Dim cmprName As String
Dim osver As OSVERSIONINFO

'注?：取得Computer Name
cmprName = String(255, 0)
len5 = 256
aa = GetComputerName(cmprName, len5)
cmprName = Left(cmprName, InStr(1, cmprName, Chr(0)) - 1)
'Debug.Print "Computer Name = "; cmprName
NameofPC = cmprName

ProgramName = HostForm.Combo1.Text & HostForm.Combo2.Text
ProgramRevisionCode = Right(HostForm.Caption, 8)



Combo1.AddItem "FT2"
Combo1.AddItem "FT3"
Combo1.AddItem "FT4"
Combo1.AddItem "===="
Combo1.AddItem "RT1"
Combo1.AddItem "RT2"
Combo1.AddItem "RT3"
Combo1.AddItem "RT4"
Combo1.AddItem "RT5"
Combo1.AddItem "RT6"
Combo1.AddItem "RT7"
Combo1.AddItem "RT8"
Combo1.AddItem "===="
Combo1.AddItem "ST1"
Combo1.AddItem "ST2"
Combo1.AddItem "ST3"
Combo1.AddItem "ST4"
Combo1.AddItem "ST5"
Combo1.AddItem "ST6"
Combo1.AddItem "ST7"
Combo1.AddItem "ST8"
Combo1.AddItem "ST9"
Combo1.AddItem "STA"
Combo1.AddItem "STB"
Combo1.AddItem "STC"

 


'注?：取得OS的版本
'osver.dwOSVersionInfoSize = Len(osver)
'aa = GetVersionEx(osver)
'Debug.Print "MajorVersion "; osver.dwMajorVersion
'Debug.Print "MinorVersion "; osver.dwMinorVersion
'Select Case osver.dwPlatformId
'Case 0
'   Debug.Print "Window 3.1"
'Case 1
'   Debug.Print "Win95"
'Case 2
'   Debug.Print "WinNT"
'End Select



End Sub


Private Sub HandlerIDText_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyReturn) Then
        OperatorNameText.SetFocus
    End If
End Sub


Private Sub LotIDText_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyReturn) Then
        HandlerIDText.SetFocus
    End If
End Sub

Private Sub MakeSure_Click()





If RunCardNOText.Text = "" _
      Or LotIDText.Text = "" _
      Or DeviceIDText.Text = "" _
       Or HandlerIDText.Text = "" _
         Or Combo1.Text = "" _
      Or OperatorNameText.Text = "" Then
    
    MsgBox "不可有空白欄"
    Exit Sub
Else
    DeviceID = DeviceIDText.Text
    RunCardNO = RunCardNOText.Text
    LotID = LotIDText.Text
    HandlerID = HandlerIDText.Text
    OperatorName = OperatorNameText.Text
    
    ProcessID = Combo1.Text
   ' counter initial
    Bin1Site1 = 0
    Bin2Site1 = 0
    Bin3Site1 = 0
    Bin4Site1 = 0
    Bin5Site1 = 0
    Bin1Site2 = 0
    Bin2Site2 = 0
    Bin3Site2 = 0
    Bin4Site2 = 0
    Bin5Site2 = 0
    
    ' time initail
    Dim StartDay As String
    Dim StartSecond As String
    Dim SNow As String
    StartSecond = Format(Now, "HH:MM:SS")
    StartDay = Format(Now, "YYYY/MM/DD")
    
    
    StartAt = StartDay & Space(1) & StartSecond
    
    ReportBegin = 1

    InsertDB
    Me.Hide
End If
    
    
       

End Sub

'  Or LotIDText.Text = "" _
'    Or DeviceIDText.Text = "" _
'

Private Sub RunCardNOText_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyReturn) Then
        LotIDText.SetFocus
    End If
End Sub
