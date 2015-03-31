VERSION 5.00
Begin VB.Form ReportForm 
   Caption         =   "報表設定"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10785
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
   ScaleWidth      =   10785
   StartUpPosition =   3  '系統預設值
   Begin VB.ComboBox ProcessIDCombo 
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
      Width           =   5535
   End
   Begin VB.CommandButton MakeSure 
      Caption         =   "確定"
      Height          =   615
      Left            =   3600
      TabIndex        =   10
      Top             =   5880
      Width           =   2415
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
      Width           =   5415
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
      Width           =   5415
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
      Width           =   5295
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
      Width           =   5295
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
      Width           =   5295
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
'On Error Resume Next
Dim oDB As ADOX.Catalog
Dim sDBPAth As String
Dim sConStr As String
Dim oCn As ADODB.Connection
Dim oCM As ADODB.Command
Dim RS As ADODB.Recordset

    '--------------------------------
    ' set Path and connection string
    '--------------------------------
    If No8PCard Then
        sDBPAth = "D:\SLT Summary\Summary.mdb"
        If Dir(sDBPAth, vbNormal + vbDirectory) = " " Then
            MsgBox "MDB no EXIST"
            Exit Sub
        End If
        
        EndAt = "NA"
        sConStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & "D:\SLT Summary" & "\SLT.mdb"
    Else
        sDBPAth = "D:\SLT Summary\MultiSummary.mdb"
        If Dir(sDBPAth, vbNormal + vbDirectory) = " " Then
            MsgBox "MDB no EXIST"
            Exit Sub
        End If
        
        EndAt = "NA"
        sConStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & "D:\SLT Summary" & "\MultiSLT.mdb"
    End If
    
    ' ------------------------
    ' Create New ADOX Object
    ' ------------------------
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

    If No8PCard Then
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
    Else
        oCM.CommandText = "INSERT INTO Summary VALUES('" & NameofPC & "','" & ProgramName & "','" & ProgramRevisionCode & "','" & DeviceID & "','" & RunCardNO & "','" & ProcessID & "','" & LotID & " ','" & StartAt & "','" & EndAt & "','" & HandlerID & "','" & OperatorName & "','" & Sites & "'," _
        & Bin1Counter(0) & "," & Bin1Counter(1) & "," & Bin1Counter(2) & "," & Bin1Counter(3) & "," & Bin1Counter(4) & "," & Bin1Counter(5) & "," & Bin1Counter(6) & "," & Bin1Counter(7) & "," _
        & Bin2Counter(0) & "," & Bin2Counter(1) & "," & Bin2Counter(2) & "," & Bin2Counter(3) & "," & Bin2Counter(4) & "," & Bin2Counter(5) & "," & Bin2Counter(6) & "," & Bin2Counter(7) & "," _
        & Bin3Counter(0) & "," & Bin3Counter(1) & "," & Bin3Counter(2) & "," & Bin3Counter(3) & "," & Bin3Counter(4) & "," & Bin3Counter(5) & "," & Bin3Counter(6) & "," & Bin3Counter(7) & "," _
        & Bin4Counter(0) & "," & Bin4Counter(1) & "," & Bin4Counter(2) & "," & Bin4Counter(3) & "," & Bin4Counter(4) & "," & Bin4Counter(5) & "," & Bin4Counter(6) & "," & Bin4Counter(7) & "," _
        & Bin5Counter(0) & "," & Bin5Counter(1) & "," & Bin5Counter(2) & "," & Bin5Counter(3) & "," & Bin5Counter(4) & "," & Bin5Counter(5) & "," & Bin5Counter(6) & "," & Bin5Counter(7) & ")"
    End If
    
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
    
    If No8PCard Then
        If Dir(sDBPAth & "\SLT.mdb", vbNormal + vbDirectory) <> "" Then
            Exit Sub
        End If
     
        sConStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & "D:\SLT Summary" & "\SLT.mdb"
    Else
    
        If Dir(sDBPAth & "\MultiSLT.mdb", vbNormal + vbDirectory) <> "" Then
           Exit Sub
        End If
    
        sConStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & "D:\SLT Summary" & "\MultiSLT.mdb"
    End If
     
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
    
    If No8PCard Then
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
    Else
        oCM.CommandText = "Create Table Summary ([NameofPC] Text(20), [ProgramName] Text(20), [ProgramRevisionCode] Text(20), [DeviceID] Text(50), [RunCardNO] Text(20), [ProcessID] Text(20), [LotID] Text(20), [StartAt] Text(20), [EndAt] Text(20), [HandlerID] Text(20), [OperatorName] Text(20),[Sites] Text(2)," _
            & "[Bin1_0] Int,[Bin1_1] Int,[Bin1_2] Int,[Bin1_3] Int,[Bin1_4] Int,[Bin1_5] Int,[Bin1_6] Int,[Bin1_7] Int," _
            & "[Bin2_0] Int,[Bin2_1] Int,[Bin2_2] Int,[Bin2_3] Int,[Bin2_4] Int,[Bin2_5] Int,[Bin2_6] Int,[Bin2_7] Int ," _
            & "[Bin3_0] Int,[Bin3_1] Int,[Bin3_2] Int,[Bin3_3] Int,[Bin3_4] Int,[Bin3_5] Int,[Bin3_6] Int,[Bin3_7] Int ," _
            & "[Bin4_0] Int,[Bin4_1] Int,[Bin4_2] Int,[Bin4_3] Int,[Bin4_4] Int,[Bin4_5] Int,[Bin4_6] Int,[Bin4_7] Int ," _
            & "[Bin5_0] Int,[Bin5_1] Int,[Bin5_2] Int,[Bin5_3] Int,[Bin5_4] Int,[Bin5_5] Int,[Bin5_6] Int,[Bin5_7] Int)"
    End If
    
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
Dim len5 As Long, aa As Long
Dim cmprName As String
Dim osver As OSVERSIONINFO

    If ReportDebug = 1 Then
        DeviceIDText.Text = "AU6471b51-JFL-GR"
        RunCardNOText.Text = "2008CLOP0040"
        LotIDText.Text = "A697940.1"
        HandlerIDText.Text = "NS-1"
        OperatorNameText.Text = "149"
    End If

    Call CreateDB

    '注?：取得Computer Name
    cmprName = String(255, 0)
    len5 = 256
    aa = GetComputerName(cmprName, len5)
    cmprName = Left(cmprName, InStr(1, cmprName, Chr(0)) - 1)
    'Debug.Print "Computer Name = "; cmprName
    NameofPC = cmprName
    
    ProgramName = Trim(Host.ChipNameCombo.Text) & Trim(Host.ChipNameCombo2.Text)
    ProgramRevisionCode = Right(Host.Caption, 8)
    
    ProcessIDCombo.AddItem "FT2"
    ProcessIDCombo.AddItem "FT3"
    ProcessIDCombo.AddItem "FT4"
    ProcessIDCombo.AddItem "===="
    ProcessIDCombo.AddItem "RT1"
    ProcessIDCombo.AddItem "RT2"
    ProcessIDCombo.AddItem "RT3"
    ProcessIDCombo.AddItem "RT4"
    ProcessIDCombo.AddItem "RT5"
    ProcessIDCombo.AddItem "RT6"
    ProcessIDCombo.AddItem "RT7"
    ProcessIDCombo.AddItem "RT8"
    ProcessIDCombo.AddItem "===="
    ProcessIDCombo.AddItem "ST1"
    ProcessIDCombo.AddItem "ST2"
    ProcessIDCombo.AddItem "ST3"
    ProcessIDCombo.AddItem "ST4"
    ProcessIDCombo.AddItem "ST5"
    ProcessIDCombo.AddItem "ST6"
    ProcessIDCombo.AddItem "ST7"
    ProcessIDCombo.AddItem "ST8"
    ProcessIDCombo.AddItem "ST9"
    ProcessIDCombo.AddItem "STA"
    ProcessIDCombo.AddItem "STB"
    ProcessIDCombo.AddItem "STC"

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

Dim i As Byte

    If RunCardNOText.Text = "" _
            Or LotIDText.Text = "" _
            Or DeviceIDText.Text = "" _
            Or HandlerIDText.Text = "" _
            Or ProcessIDCombo.Text = "" _
            Or OperatorNameText.Text = "" Then
        
        MsgBox "不可有空白欄"
        Exit Sub
    Else
        DeviceID = DeviceIDText.Text
        RunCardNO = RunCardNOText.Text
        LotID = LotIDText.Text
        HandlerID = HandlerIDText.Text
        OperatorName = OperatorNameText.Text
        
        ProcessID = ProcessIDCombo.Text
        Sites = Host.SiteCombo.Text
       ' counter initial
       
        For i = 0 To 7
            Bin1Counter(i) = 0
            Bin2Counter(i) = 0
            Bin3Counter(i) = 0
            Bin4Counter(i) = 0
            Bin5Counter(i) = 0
        Next
        
        ' time initail
        Dim StartDay As String
        Dim StartSecond As String
        Dim SNow As String
        StartSecond = Format(Now, "HH:MM:SS")
        StartDay = Format(Now, "YYYY/MM/DD")
        
        StartAt = StartDay & Space(1) & StartSecond
        EndAt = EndDay & Space(1) & EndSecond
        
        ReportBegin = 1
    
        InsertDB
        Me.Hide
    End If
    
End Sub

Private Sub RunCardNOText_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyReturn) Then
        LotIDText.SetFocus
    End If
End Sub
