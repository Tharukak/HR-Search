VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEMF 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Employee Master File"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9825
   BeginProperty Font 
      Name            =   "Calibri"
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
   ScaleHeight     =   8760
   ScaleWidth      =   9825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   7575
      Left            =   120
      TabIndex        =   19
      Top             =   840
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   13361
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Employee Details"
      TabPicture(0)   =   "frmEMF.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Employee &Inquery"
      TabPicture(1)   =   "frmEMF.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   7095
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   8535
         Begin VB.CheckBox chkEsc 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Special Approval Authority"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   5160
            TabIndex        =   42
            Top             =   3000
            Width           =   2175
         End
         Begin VB.CommandButton cmdUpdate 
            Caption         =   "&Update"
            Enabled         =   0   'False
            Height          =   615
            Left            =   7680
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Update Record to HR Search"
            Top             =   5400
            Width           =   615
         End
         Begin VB.TextBox txtUDName 
            Height          =   285
            Left            =   5040
            TabIndex        =   4
            Top             =   720
            Width           =   2415
         End
         Begin VB.CommandButton Command3 
            Height          =   255
            Left            =   6960
            TabIndex        =   38
            Top             =   1440
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   5400
            TabIndex        =   6
            Top             =   1440
            Width           =   1455
         End
         Begin VB.ComboBox cmbDivisionName 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   1440
            Width           =   3495
         End
         Begin VB.Frame Frame3 
            Height          =   3255
            Left            =   7560
            TabIndex        =   37
            Top             =   3720
            Width           =   855
            Begin VB.CommandButton cmdEdit 
               Caption         =   "&Edit"
               Enabled         =   0   'False
               Height          =   615
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   40
               ToolTipText     =   "Edit Record from HR Search"
               Top             =   1080
               Width           =   615
            End
            Begin VB.CommandButton cmdDelete 
               Caption         =   "&Delete"
               Height          =   735
               Left            =   120
               Picture         =   "frmEMF.frx":0038
               Style           =   1  'Graphical
               TabIndex        =   16
               ToolTipText     =   "Delete Record to HR Search"
               Top             =   2400
               Width           =   615
            End
            Begin VB.CommandButton cmdAdd 
               Caption         =   "&Add"
               Height          =   735
               Left            =   120
               Picture         =   "frmEMF.frx":053E
               Style           =   1  'Graphical
               TabIndex        =   15
               ToolTipText     =   "Add Record to HR Search"
               Top             =   240
               Width           =   615
            End
         End
         Begin VB.TextBox txtUEmpNo 
            Height          =   285
            Left            =   1800
            TabIndex        =   2
            Top             =   720
            Width           =   1455
         End
         Begin VB.ComboBox cmbCostCentre 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox txtUCont1 
            Height          =   285
            Left            =   1800
            TabIndex        =   7
            Top             =   1800
            Width           =   1935
         End
         Begin VB.TextBox txtUCont2 
            Height          =   285
            Left            =   5400
            TabIndex        =   8
            Top             =   1800
            Width           =   2055
         End
         Begin VB.TextBox txtUeMail 
            Height          =   285
            Left            =   1800
            TabIndex        =   9
            Top             =   2160
            Width           =   5655
         End
         Begin VB.TextBox txtREmpNo 
            Height          =   285
            Left            =   1800
            TabIndex        =   10
            Top             =   3120
            Width           =   1455
         End
         Begin VB.CommandButton cmdEmp 
            Height          =   255
            Left            =   3360
            TabIndex        =   3
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtReMail 
            Height          =   285
            Left            =   1800
            TabIndex        =   14
            Top             =   4200
            Width           =   5655
         End
         Begin VB.TextBox txtRCont1 
            Height          =   285
            Left            =   1800
            TabIndex        =   12
            Top             =   3840
            Width           =   1935
         End
         Begin VB.TextBox txtRCont2 
            Height          =   285
            Left            =   5400
            TabIndex        =   13
            Top             =   3840
            Width           =   2055
         End
         Begin VB.CommandButton cmdREmp 
            Height          =   255
            Left            =   3360
            TabIndex        =   11
            Top             =   3120
            Width           =   495
         End
         Begin MSFlexGridLib.MSFlexGrid MSFEmployee 
            Height          =   2415
            Left            =   120
            TabIndex        =   21
            Top             =   4560
            Width           =   7335
            _ExtentX        =   12938
            _ExtentY        =   4260
            _Version        =   393216
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Display Name"
            Height          =   255
            Left            =   3960
            TabIndex        =   39
            Top             =   720
            Width           =   975
         End
         Begin VB.Image Image4 
            Height          =   1095
            Left            =   7560
            Picture         =   "frmEMF.frx":0B45
            Stretch         =   -1  'True
            Top             =   2640
            Width           =   855
         End
         Begin VB.Image Image3 
            Height          =   1095
            Left            =   7560
            Picture         =   "frmEMF.frx":2681
            Stretch         =   -1  'True
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Employee No."
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Cost Centre"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Full Name"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Division Name"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Label lblUName 
            BackColor       =   &H00E0E0E0&
            Height          =   255
            Left            =   1800
            TabIndex        =   32
            Top             =   1080
            Width           =   5655
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Contact Number 01"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   1800
            Width           =   1575
         End
         Begin VB.Label Label7 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Contact Number 02"
            Height          =   255
            Left            =   3840
            TabIndex        =   30
            Top             =   1800
            Width           =   1575
         End
         Begin VB.Label Label8 
            BackColor       =   &H00FFFFFF&
            Caption         =   "e-Mail "
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   2160
            Width           =   1575
         End
         Begin VB.Label Label9 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Reporting Emp. No."
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   3120
            Width           =   1575
         End
         Begin VB.Label Label10 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Reporting e-Mail"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   4200
            Width           =   1575
         End
         Begin VB.Label Label11 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Contact Number 01"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   3840
            Width           =   1575
         End
         Begin VB.Label Label12 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Contact Number 02"
            Height          =   255
            Left            =   3840
            TabIndex        =   25
            Top             =   3840
            Width           =   1575
         End
         Begin VB.Label Label13 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Reporting Name"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   3480
            Width           =   1575
         End
         Begin VB.Label lblRName 
            BackColor       =   &H00E0E0E0&
            Height          =   255
            Left            =   1800
            TabIndex        =   23
            Top             =   3480
            Width           =   5655
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "Reporting Person Details"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   2640
            Width           =   7335
         End
      End
   End
   Begin VB.Frame Frame2 
      Height          =   7770
      Left            =   9000
      TabIndex        =   17
      Top             =   720
      Width           =   735
      Begin VB.CommandButton cmdExit 
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
         Left            =   120
         Picture         =   "frmEMF.frx":3D27
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Exit from the System"
         Top             =   7080
         Width           =   495
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   8505
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   10
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "10/17/2020"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "3:07 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Enabled         =   0   'False
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   2280
      Picture         =   "frmEMF.frx":456C
      Stretch         =   -1  'True
      Top             =   120
      Width           =   5175
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   14  'Copy Pen
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   2160
      Top             =   0
      Width           =   11175
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   0
      Picture         =   "frmEMF.frx":5AF5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "frmEMF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DEmpNo, UEmpNo, REmpNo, UFName, RFName, UDName, Ucon1, Ucon2, Uemail, RCon1, RCon2, ReMail As String
Dim SP_Approve As Integer

Private Sub cmbCostCentre_Click()
On Error GoTo er_EH:
    If cmdUpdate.Enabled = False Then
        cmbDivisionName.Clear
        PR_HRS_Open_CON
        Set RS = New ADODB.Recordset
        RS.Open "SELECT Distinct D_Name FROM HRS_sys_Division Where Com_Code='" & Trim(cmbCostCentre.Text) & "' Order by D_Name", HRS, adOpenStatic, adLockReadOnly
        Do While RS.EOF = False
            cmbDivisionName.AddItem Trim(RS!D_Name)
            RS.MoveNext
        Loop
        cmbDivisionName.AddItem "-Select-"
        cmbDivisionName.Text = "-Select-"
    End If
    PR_HRS_Close_CON
    Exit Sub
er_EH:
    MsgBox Err.Description, vbCritical
    Err.Clear
End Sub

Private Sub cmdAdd_Click()
On Error GoTo er_EH:
    If Trim(cmbCostCentre.Text) = "" Then
        MsgBox "Cost Centre NOT Selected", vbExclamation
        cmbCostCentre.SetFocus
        Exit Sub
    Else
        Com_Code = Trim(cmbCostCentre.Text)
    End If
    
    If Trim(txtUEmpNo.Text) = "" Then
        MsgBox "Employee Number NOT Entered", vbExclamation
        txtUEmpNo.SetFocus
        Exit Sub
    End If
    
    If Trim(cmbDivisionName.Text) = "" Then
        MsgBox "Division Name NOT Selected", vbExclamation
        cmbDivisionName.SetFocus
        Exit Sub
    Else
        Division_Name = Trim(cmbDivisionName.Text)
        FN_DivisionID Division_Name, Com_Code
    End If
    
    If Trim(txtUCont1.Text) = "" Then
        MsgBox "Contact Number 1 NOT Entered", vbExclamation
        txtUCont1.SetFocus
        Exit Sub
    End If
    
    If Trim(txtUCont2.Text) = "" Then
        MsgBox "Contact Number 2 NOT Entered", vbExclamation
        txtUCont2.SetFocus
        Exit Sub
    End If
    
    If Trim(txtUDName.Text) = "" Then
        MsgBox "User Display Name NOT Entered", vbExclamation
        txtUDName.SetFocus
        Exit Sub
    End If
    
    If Trim(txtREmpNo.Text) = "" Then
        MsgBox "User e-Mail NOT Entered", vbExclamation
        txtREmpNo.SetFocus
        Exit Sub
    End If
    
    If Trim(txtREmpNo.Text) = "" Then
        MsgBox "Reporting Person Employee Number NOT Entered", vbExclamation
        txtREmpNo.SetFocus
        Exit Sub
    End If
    
    If Trim(txtRCont1.Text) = "" Then
        MsgBox "Reporting Person Contact Number 1 NOT Entered", vbExclamation
        txtRCont1.SetFocus
        Exit Sub
    End If
    
    If Trim(txtRCont2.Text) = "" Then
        MsgBox "Reporting Person Contact Number 2 NOT Entered", vbExclamation
        txtRCont2.SetFocus
        Exit Sub
    End If
    
    If Trim(txtReMail.Text) = "" Then
        MsgBox "Reporting Person e-Mail NOT Entered", vbExclamation
        txtReMail.SetFocus
        Exit Sub
    End If
    
    If chkEsc.Value = 1 Then
        SP_Approve = 1
    Else
        SP_Approve = 0
    End If
    
    
    UEmpNo = Trim(txtUEmpNo.Text)
    UDName = Trim(txtUDName.Text)
    UFName = Trim(lblUName.Caption)
    Ucon1 = Trim(txtUCont1.Text)
    Ucon2 = Trim(txtUCont2.Text)
    Uemail = Trim(txtUeMail.Text)
    REmpNo = Trim(txtREmpNo.Text)
    RFName = Trim(lblRName.Caption)
    RCon1 = Trim(txtRCont1.Text)
    RCon2 = Trim(txtRCont2.Text)
    ReMail = Trim(txtReMail.Text)

    PR_HRS_Open_CON
    Set DupRS = New ADODB.Recordset
    DupRS.Open "Select * from HRS_sys_MSTR_Employee Where U_ID='" & UEmpNo & "'", HRS, adOpenStatic, adLockReadOnly
    If DupRS.EOF = False Then
        MsgBox "Employee Already Exist", vbExclamation
    Else
        HRS.Execute "INSERT INTO HRS_sys_MSTR_Employee(U_ID,U_FName,U_DName,U_Contact1,U_Contact2,U_eMail,R_ID,R_eMail,Com_Code,D_Code,R_Contact1,R_Contact2,SP_Approve)" _
                + "Values('" & UEmpNo & "','" & UFName & "','" & UDName & "','" & Ucon1 & "','" & Ucon2 & "','" & Uemail & "','" & REmpNo & "','" & ReMail & "','" & Com_Code & "','" & Division_Code & "','" & RCon1 & "','" & RCon2 & "'," & SP_Approve & ")"
        
        Set DupRS = New ADODB.Recordset
        DupRS.Open "Select * from HRS_sys_Login Where U_ID='" & UEmpNo & "'", HRS, adOpenStatic, adLockReadOnly
        If DupRS.EOF = True Then
            HRS.Execute "INSERT INTO HRS_sys_Login(U_ID,U_PW) Values('" & UEmpNo & "','welcome')"
            FN_User_Profile Trim(txtUEmpNo.Text)
        Else
            MsgBox "This Employee already system user", vbExclamation
        End If
        MsgBox "Employee Registered Successfully", vbInformation
    End If
    PR_Grid
    PR_HRS_Close_CON
    Exit Sub
    
er_EH:
    MsgBox Err.Description, vbCritical
    Err.Clear
End Sub

Private Sub cmdDelete_Click()
On Error GoTo er_EH:
    If MsgBox("Are you sure Do you Want to Delete your selected Employee ?", vbQuestion + vbYesNo, "Delete Record") = vbYes Then
        PR_HRS_Open_CON
        HRS.Execute "Delete from HRS_sys_MSTR_Employee Where U_ID='" & DEmpNo & "'"
        HRS.Execute "Delete from HRS_sys_Login Where U_ID='" & DEmpNo & "'"
        PR_HRS_Close_CON
        PR_Grid
    End If
    Exit Sub
er_EH:
    MsgBox Err.Description, vbCritical
    Err.Clear
End Sub

Private Sub cmdEdit_Click()
    cmbCostCentre.Enabled = True
    txtUEmpNo.Enabled = True
    cmbDivisionName.Enabled = True
    txtUCont1.Enabled = True
    txtUCont2.Enabled = True
    txtUDName.Enabled = True
    txtREmpNo.Enabled = True
    txtREmpNo.Enabled = True
    txtRCont1.Enabled = True
    txtRCont2.Enabled = True
    txtReMail.Enabled = True
    cmdEdit.Enabled = False
End Sub

Private Sub cmdEmp_Click()
On Error GoTo er_EH:
    If Trim(txtUEmpNo.Text) = "" Then
        MsgBox "Employee Number NOT Entered", vbExclamation
        txtUEmpNo.SetFocus
    Else
        UEmpNo = Trim(txtUEmpNo.Text)
        PR_Find_Employee
        lblUName.Caption = UFName
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
Private Sub cmdREmp_Click()
On erro GoTo er_EH:
    If Trim(txtREmpNo.Text) = "" Then
        MsgBox "Reporting Person Employee Number NOT Entered", vbExclamation
        txtREmpNo.SetFocus
    Else
        UEmpNo = Trim(txtREmpNo.Text)
        PR_Find_Employee
        lblRName.Caption = UFName
    End If
    Exit Sub
    
er_EH:
    MsgBox Err.Description, vbCritical
    Err.Clear
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo er_EH:
    PR_HRS_Open_CON
    Set RS = New ADODB.Recordset
    RS.Open "Select * from HRS_sys_MSTR_Employee Where U_ID='" & Trim(txtUEmpNo.Text) & "'", HRS, adOpenStatic, adLockPessimistic
    If RS.EOF = False Then
        Com_Code = Trim(cmbCostCentre.Text)
        RS.Fields("Com_Code") = Trim(cmbCostCentre.Text)
        Division_Name = Trim(cmbDivisionName.Text)
        FN_DivisionID Division_Name, Com_Code
        RS.Fields("D_Code") = Division_Code
        RS.Fields("U_Contact1") = Trim(txtUCont1.Text)
        RS.Fields("U_Contact2") = Trim(txtUCont2.Text)
        RS.Fields("U_DName") = Trim(txtUDName.Text)
        RS.Fields("U_email") = Trim(txtUeMail.Text)
        RS.Fields("R_ID") = Trim(txtREmpNo.Text)
        RS.Fields("R_Contact1") = Trim(txtRCont1.Text)
        RS.Fields("R_Contact2") = Trim(txtRCont2.Text)
        RS.Fields("R_Email") = Trim(txtReMail.Text)
        If chkEsc.Value = 1 Then
            SP_Approve = 1
        Else
            SP_Approve = 0
        End If
        RS.Fields("SP_Approve") = SP_Approve
        RS.Update
        PR_Grid
        MsgBox "Employee Updated Successfully", vbInformation
        cmdUpdate.Enabled = False
        cmdAdd.Enabled = True
        cmdDelete.Enabled = True
        PR_Object_Clear
    End If
    PR_HRS_Close_CON
    Exit Sub
er_EH:
    MsgBox Err.Description, vbCritical
    Err.Clear
    
End Sub

Private Sub Form_Load()
On Error GoTo er_EH:
    cmbCostCentre.Clear
    Set RS = New ADODB.Recordset
    PR_HRS_Open_CON
    RS.Open "Select Com_Code from HRS_sys_Company Order by Com_Code", HRS, adOpenStatic, adLockReadOnly
    Do While RS.EOF = False
        cmbCostCentre.AddItem Trim(RS!Com_Code)
        RS.MoveNext
    Loop
    cmbCostCentre.AddItem "-Select-"
    cmbCostCentre.Text = "-Select-"
    PR_Grid
    PR_HRS_Close_CON
    Exit Sub
er_EH:
    MsgBox Err.Description, vbCritical
    Err.Clear
End Sub

Public Sub PR_Find_Employee()
On Error GoTo er_EH:
    PR_HRS_Open_CON
    Set RS = New ADODB.Recordset
    RS.Open "Select * from HCMV_sys_MSTR_Employee Where EMP_Number='" & UEmpNo & "'", HRS, adOpenStatic, adLockReadOnly
    If RS.EOF = False Then
        UFName = Trim(RS!Emp_FullName)
    Else
        UFName = ""
        MsgBox "Employee Number NOT Found", vbExclamation
    End If
    PR_HRS_Close_CON
    Exit Sub
er_EH:
    MsgBox Err.Description, vbCritical
    Err.Clear
End Sub

Public Sub PR_Grid()
    On Error GoTo er_EH:
    MSFEmployee.Clear
    PR_HRS_Open_CON
    R = 1
    MSFEmployee.Cols = 5
    Set GRDRS = New ADODB.Recordset
    GRDRS.Open "Select * from HRSV_sys_MSTR_Employee Order by U_ID", HRS, adOpenStatic, adLockReadOnly
    Do While GRDRS.EOF = False
        MSFEmployee.Rows = R + 1
        MSFEmployee.TextMatrix(R, 0) = Trim(GRDRS!U_ID)
        MSFEmployee.TextMatrix(R, 1) = Trim(GRDRS!U_DName)
        MSFEmployee.TextMatrix(R, 2) = Trim(GRDRS!U_email)
        MSFEmployee.TextMatrix(R, 3) = Trim(GRDRS!R_Name)
        MSFEmployee.TextMatrix(R, 4) = Trim(GRDRS!R_Email)
        R = R + 1
        GRDRS.MoveNext
    Loop
    MSFEmployee.TextMatrix(0, 0) = "Emp No."
    MSFEmployee.TextMatrix(0, 1) = "Emp Display Name"
    MSFEmployee.TextMatrix(0, 2) = "Emp e-Mail"
    MSFEmployee.TextMatrix(0, 3) = "Reporting Peron"
    MSFEmployee.TextMatrix(0, 4) = "Reporting e-Mail"
    MSFEmployee.ColWidth(0) = 1000
    MSFEmployee.ColWidth(1) = 3000
    MSFEmployee.ColWidth(2) = 2000
    MSFEmployee.ColWidth(3) = 3000
    MSFEmployee.ColWidth(4) = 2000
    PR_HRS_Close_CON
    Exit Sub
er_EH:
    MsgBox Err.Description
    Err.Clear
End Sub

Private Sub MSFEmployee_Click()
    DEmpNo = MSFEmployee.TextMatrix(MSFEmployee.Row, 0)
End Sub

Public Sub PR_Object_Clear()
    cmbCostCentre.Text = "-Select-"
    txtUEmpNo.Text = ""
    cmbDivisionName.Text = "-Select-"
    txtUCont1.Text = ""
    txtUCont2.Text = ""
    txtUDName.Text = ""
    txtREmpNo.Text = ""
    txtREmpNo.Text = ""
    txtRCont1.Text = ""
    txtRCont2.Text = ""
    txtReMail.Text = ""
    txtUeMail.Text = ""
    lblUName.Caption = ""
    lblRName.Caption = ""
    chkEsc.Value = 0
End Sub

Public Sub PR_Edit()
On Error GoTo er_EH:
    PR_HRS_Open_CON
    Set GRDRS = New ADODB.Recordset
    GRDRS.Open "Select * from HRSV_sys_MSTR_Employee Where U_ID='" & DEmpNo & "'", HRS, adOpenStatic, adLockReadOnly
    If GRDRS.EOF = False Then
        cmbCostCentre.Text = Trim(GRDRS!Com_Code)
        txtUEmpNo.Text = Trim(GRDRS!U_ID)
        cmbDivisionName.Text = Trim(GRDRS!D_Name)
        txtUCont1.Text = Trim(GRDRS!U_Contact1)
        txtUCont2.Text = Trim(GRDRS!U_Contact2)
        txtUDName.Text = Trim(GRDRS!U_DName)
        txtUeMail.Text = Trim(GRDRS!U_email)
        txtREmpNo.Text = Trim(GRDRS!R_ID)
        txtRCont1.Text = Trim(GRDRS!R_Contact1)
        txtRCont2.Text = Trim(GRDRS!R_Contact1)
        txtReMail.Text = Trim(GRDRS!R_Email)
        If GRDRS!SP_Approve = True Then
            chkEsc.Value = 1
        Else
            chkEsc.Value = 0
        End If
    End If
    Call cmdEmp_Click
    Call cmdREmp_Click
    PR_HRS_Close_CON
    cmbCostCentre.Enabled = False
    txtUEmpNo.Enabled = False
    cmbDivisionName.Enabled = False
    txtUCont1.Enabled = False
    txtUCont2.Enabled = False
    txtUDName.Enabled = False
    txtREmpNo.Enabled = False
    txtREmpNo.Enabled = False
    txtRCont1.Enabled = False
    txtRCont2.Enabled = False
    txtReMail.Enabled = False
    Exit Sub
    
er_EH:
    MsgBox Err.Description, vbCritical
    Err.Clear
End Sub

Private Sub MSFEmployee_DblClick()
    PR_Edit
    cmdAdd.Enabled = False
    cmdDelete.Enabled = False
    cmdEdit.Enabled = True
    cmdUpdate.Enabled = True
End Sub

Public Sub FN_User_Profile(ByRef New_User_ID As String)
    Dim cmd As ADODB.Command
    Dim RS1 As New ADODB.Recordset
    Set RS1 = New ADODB.Recordset
    Set cmd = New ADODB.Command
    Dim strconnect As String

    PR_HRS_Open_CON
    cmd.ActiveConnection = HRS
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "HRS_SP_Profile"
    cmd.Parameters.Append cmd.CreateParameter("@User_ID", adChar, adParamInput, 10, New_User_ID)
    cmd.CommandTimeout = 0
    Set RS1 = cmd.Execute
    Set cmd.ActiveConnection = Nothing
End Sub
