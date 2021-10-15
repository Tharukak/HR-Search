VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmLogin 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Welcome to HR Search"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5520
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   0
      Picture         =   "frmLogin.frx":4B85A
      ScaleHeight     =   795
      ScaleWidth      =   2355
      TabIndex        =   8
      Top             =   0
      Width           =   2415
   End
   Begin VB.TextBox txtPW 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   2
      ToolTipText     =   "Enter Your Password"
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox txtUID 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      ToolTipText     =   "Enter your Global Employee Number"
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "&Login"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   1080
      Width           =   1695
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   2175
      Width           =   5520
      _ExtentX        =   9737
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7408
            MinWidth        =   7408
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   0
      TabIndex        =   5
      Top             =   840
      Width           =   5535
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "User ID"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Neuropolitical Rg"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   495
      Left            =   2520
      TabIndex        =   9
      Top             =   180
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   800
      Left            =   2400
      Top             =   0
      Width           =   3135
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    Close All
    Unload Me
End Sub

Private Sub cmdLogin_Click()
    If Trim(txtUID.Text) = "" Then
        MsgBox "User ID NOT Found", vbExclamation
        Exit Sub
    Else
        UserID = Trim(txtUID.Text)
    End If
    
    If Trim(txtPW.Text) = "" Then
        MsgBox "Password NOT Found", vbExclamation
        Exit Sub
    Else
        PW = Trim(txtPW.Text)
    End If
    PR_Login
End Sub

Private Sub Form_Load()
    db_Connect
End Sub

Public Sub PR_Login()
    On Error GoTo er_EH:
    Set RS = New ADODB.Recordset
    PR_HRS_Open_CON
    RS.Open "Select * from HRS_sys_Login Where U_ID='" & UserID & "'", HRS, adOpenStatic, adLockReadOnly
    If RS.EOF = False Then
        If PW = Trim(RS!U_PW) Then
            Dim RS2 As ADODB.Recordset
            Set RS2 = New ADODB.Recordset
            RS2.Open "Select * From HRS_sys_MSTR_Employee Where U_ID='" & UserID & "'", HRS, adOpenStatic, adLockReadOnly
            If RS.EOF = False Then
                MDIHome.StatusBar1.Panels(1).Text = Trim(RS2!U_DName)
                Medi_U_Name = Trim(RS2!U_DName)
                U_Com_Code = Trim(RS2!Com_Code)
                U_D_Code = Val(RS2!D_Code)
                LOG_REGISTER
                MDIHome.Show
                Unload frmLogin
            Else
                MsgBox "Employee Master File NOT created Properly", vbExclamation
                Exit Sub
            End If
        Else
            MsgBox "Invalid Password", vbExclamation
        End If
    Else
        MsgBox "User ID NOT Found", vbExclamation
    End If
    PR_HRS_Close_CON
    Exit Sub
    
er_EH:
    MsgBox Err.Description, vbCritical
    Err.Clear
End Sub

Public Sub LOG_REGISTER()
    If HRS.State = adStateClosed Then HRS.Open
    HRS.Execute "INSERT INTO HRS_sys_Log(TRANS_DATE_TIME,MECHINENAME,Major_Version,Minor_Version,Revision_Version,LoginName) VALUES('" & Format(Date, "MM/dd/yyyy") + " " + Format(Time, "hh:mm:ss") & "','" & VBA.Environ("COMPUTERNAME") & "'," & App.Major & "," & App.Minor & "," & App.Revision & ",'" & UCase(Environ("USERNAME")) & "')"
End Sub

Private Sub txtPW_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdLogin_Click
    End If
End Sub

Private Sub txtUID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPW.SetFocus
    End If
End Sub
