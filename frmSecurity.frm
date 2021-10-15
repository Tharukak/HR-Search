VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmSecurity 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Security Module"
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9600
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
   ScaleHeight     =   8460
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdProfile 
      Caption         =   "&Create User Profile"
      Height          =   375
      Left            =   6600
      TabIndex        =   9
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find Employee"
      Height          =   375
      Left            =   4560
      TabIndex        =   8
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox txtUID 
      Height          =   285
      Left            =   2400
      TabIndex        =   7
      Top             =   960
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   0
      Picture         =   "frmSecurity.frx":0000
      ScaleHeight     =   795
      ScaleWidth      =   2355
      TabIndex        =   4
      Top             =   0
      Width           =   2415
   End
   Begin VB.Frame Frame2 
      Height          =   6855
      Left            =   8760
      TabIndex        =   2
      Top             =   1200
      Width           =   735
      Begin VB.CommandButton cmdExit 
         Height          =   495
         Left            =   120
         Picture         =   "frmSecurity.frx":10F8
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Exit from the System"
         Top             =   6240
         Width           =   495
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   8085
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin GridEX20.GridEX GridUserRights 
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   11880
      Version         =   "2.0"
      HoldSortSettings=   -1  'True
      DefaultGroupMode=   1
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      Options         =   8
      RecordsetType   =   1
      DatabaseName    =   $"frmSecurity.frx":193D
      GroupByBoxVisible=   0   'False
      DataMode        =   1
      ColumnHeaderHeight=   285
      IntProp3        =   $"frmSecurity.frx":19C9
      IntProp5        =   $"frmSecurity.frx":1A55
      ColumnsCount    =   6
      Column(1)       =   "frmSecurity.frx":1AE1
      Column(2)       =   "frmSecurity.frx":1C05
      Column(3)       =   "frmSecurity.frx":1D29
      Column(4)       =   "frmSecurity.frx":1E89
      Column(5)       =   "frmSecurity.frx":1FDD
      Column(6)       =   "frmSecurity.frx":2131
      GroupCount      =   3
      Group(1)        =   "frmSecurity.frx":22AD
      Group(2)        =   "frmSecurity.frx":2315
      Group(3)        =   "frmSecurity.frx":237D
      SortKeysCount   =   1
      SortKey(1)      =   "frmSecurity.frx":23E5
      FormatStylesCount=   5
      FormatStyle(1)  =   "frmSecurity.frx":244D
      FormatStyle(2)  =   "frmSecurity.frx":2575
      FormatStyle(3)  =   "frmSecurity.frx":2625
      FormatStyle(4)  =   "frmSecurity.frx":26D9
      FormatStyle(5)  =   "frmSecurity.frx":27B1
      ImageCount      =   0
      PrinterProperties=   "frmSecurity.frx":2869
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Employee Number"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Security Module"
      BeginProperty Font 
         Name            =   "Neuropolitical Rg"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   120
      Width           =   5655
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   795
      Left            =   2400
      Top             =   0
      Width           =   7215
   End
End
Attribute VB_Name = "frmSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim F_User_ID As String
Private Sub cmdExit_Click()
    Close All
    Unload Me
End Sub

Private Sub cmdFind_Click()
    If Trim(txtUID.Text) = "" Then
        MsgBox "Employee Number NOT Entered", vbExclamation
        Exit Sub
    Else
        GridUserRights.RecordSource = "Select * from HRSV_sys_Rights where User_ID='" & Trim(txtUID.Text) & "'"
    End If
    GridUserRights.HoldFields
    GridUserRights.Rebind
End Sub

Public Sub PR_User_Profile()
    Dim cmd As ADODB.Command
    Dim RS1 As New ADODB.Recordset
    Set RS1 = New ADODB.Recordset
    Set cmd = New ADODB.Command
    Dim strconnect As String

    PR_HRS_Open_CON
    cmd.ActiveConnection = HRS
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "HRS_SP_Profile"
    cmd.Parameters.Append cmd.CreateParameter("@User_ID", adChar, adParamInput, 10, F_User_ID)
    cmd.CommandTimeout = 0
    Set RS1 = cmd.Execute
    Set cmd.ActiveConnection = Nothing
End Sub

Private Sub cmdProfile_Click()
    F_User_ID = Trim(txtUID.Text)
    PR_User_Profile
End Sub
