VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmReports 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Reports"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8520
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   8520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport CRW 
      Left            =   1800
      Top             =   6120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Frame Frame2 
      Height          =   8295
      Left            =   7680
      TabIndex        =   2
      Top             =   0
      Width           =   735
      Begin VB.CommandButton cmdExit 
         Height          =   495
         Left            =   120
         Picture         =   "frmReports.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Exit from the System"
         Top             =   7680
         Width           =   495
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFReports 
      Height          =   8175
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   14420
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   8370
      Width           =   8520
      _ExtentX        =   15028
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    Close All
    Unload Me
End Sub

Private Sub Form_Load()
    PR_HRS_Open_CON
    Set RS = New ADODB.Recordset
    RS.Open "Select * from HRSV_sys_Reports Order by Main_Module_Name,Report_Index", HRS, adOpenStatic, adLockReadOnly
    MSFReports.Cols = 2
    R = 1
    Do While RS.EOF = False
        MSFReports.Rows = R + 1
        MSFReports.TextMatrix(R, 0) = Trim(RS!Report_Index)
        MSFReports.TextMatrix(R, 1) = Trim(RS!Report_Title)
        R = R + 1
        RS.MoveNext
    Loop
    PR_HRS_Close_CON
    
    MSFReports.ColWidth(0) = 1000
    MSFReports.ColWidth(1) = 8000
    
    MSFReports.TextMatrix(0, 0) = "INDEX NO."
    MSFReports.TextMatrix(0, 1) = "REPORT NAME"
End Sub
Private Sub MSFReports_DblClick()
On Error GoTo er_EH:
    Dim Report_Name As String
    PR_HRS_Open_CON
    Set RS = New ADODB.Recordset
    Report_Name = ""
    RS.Open "Select * from HRSV_sys_Reports Where Report_Index=" & Val(MSFReports.TextMatrix(MSFReports.Row, 0)), HRS, adOpenStatic, adLockReadOnly
    If RS.EOF = False Then
        Report_Name = Trim(RS!Report_Name)
    Else
        MsgBox "HR Search-Extream Report NOT Found", vbCritical
        Exit Sub
    End If
    PR_HRS_Close_CON
    
    PR_REPORT_PATH
    CRW.ReportFileName = Report_Path + Report_Name
    CRW.Connect = "dsn=" & Trim(RS!DNS_Name) & ";uid=" & Trim(RS!DNS_UID) & ";pwd=" & Trim(RS!DNS_PW)
    CRW.Action = 1
    Exit Sub
    
er_EH:
    MsgBox Err.Description, vbCritical
    Err.Clear
End Sub
