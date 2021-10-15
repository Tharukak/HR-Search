VERSION 5.00
Begin VB.Form frmChangePW 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Change My Password"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6210
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
   ScaleHeight     =   1605
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtEPW 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox txtNPW 
      Enabled         =   0   'False
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox txtCPW 
      Enabled         =   0   'False
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1200
      Width           =   2295
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "&Change Password"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   960
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   4200
      TabIndex        =   4
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "User Name"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Exsisting Password"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "New Password"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Confirm Password"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label lblUsername 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmChangePW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PW As String

Private Sub cmdChange_Click()
On Error GoTo er_EH:
    If Trim(txtNPW.Text) = "" Then
        MsgBox "New Password NOT Entered", vbExclamation
    Else
        If Trim(txtNPW.Text) = Trim(txtCPW.Text) Then
            PR_HRS_Open_CON
            Set RS = New ADODB.Recordset
            RS.Open "UPDATE HRS_sys_Login SET U_PW='" & Trim(txtNPW.Text) & "' Where U_ID='" & UserID & "'", HRS, adOpenStatic, adLockOptimistic
            MsgBox "Password Changed Successfully", vbInformation
            PR_HRS_Close_CON
        Else
            MsgBox "New Password and Confirmmed Password Incorrect", vbCritical
            Exit Sub
        End If
    End If
    Exit Sub
    
er_EH:
    MsgBox Err.Description, vbCritical
    Err.Clear
End Sub

Private Sub cmdExit_Click()
    Close All
    Unload Me
End Sub

Private Sub Form_Load()
    PR_HRS_Open_CON
    Set RS = New ADODB.Recordset
    RS.Open "Select * from HRSV_sys_Login Where u_id='" & UserID & "'", HRS, adOpenStatic, adLockReadOnly
    If RS.EOF = False Then
        lblUsername.Caption = Trim(RS!U_DName)
        PW = Trim(RS!U_PW)
    Else
        MsgBox "User ID Not Found", vbCritical
        Exit Sub
    End If
    PR_HRS_Close_CON
End Sub

Private Sub txtEPW_Change()
    txtNPW.Enabled = False
    txtCPW.Enabled = False
    If PW = Trim(txtEPW.Text) Then
        txtNPW.Enabled = True
        txtCPW.Enabled = True
        cmdChange.Enabled = True
        txtNPW.BackColor = &HFFFFFF
        txtCPW.BackColor = &HFFFFFF
    Else
        txtNPW.Enabled = False
        txtCPW.Enabled = False
        cmdChange.Enabled = False
        txtNPW.BackColor = &HC0C0FF
        txtCPW.BackColor = &HC0C0FF
    End If
End Sub
