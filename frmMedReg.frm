VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMedReg 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Medical Register"
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15930
   ControlBox      =   0   'False
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
   ScaleHeight     =   8775
   ScaleWidth      =   15930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   7455
      Left            =   14640
      TabIndex        =   17
      Top             =   960
      Width           =   1095
      Begin VB.CommandButton cmdProcess 
         Caption         =   "&Process"
         Enabled         =   0   'False
         Height          =   495
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdExit 
         Height          =   735
         Left            =   120
         Picture         =   "frmMedReg.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Exit from the System"
         Top             =   6600
         Width           =   855
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6855
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   12091
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Indoor Diagnosis"
      TabPicture(0)   =   "frmMedReg.frx":0845
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "MSFEmpSymptoms"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "MSFSymptoms"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtVar1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Precription"
      TabPicture(1)   =   "frmMedReg.frx":0861
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label7"
      Tab(1).Control(1)=   "Label8"
      Tab(1).Control(2)=   "Label9"
      Tab(1).Control(3)=   "Label10"
      Tab(1).Control(4)=   "MSFEmpDrugs"
      Tab(1).Control(5)=   "MSFDrugs"
      Tab(1).Control(6)=   "Frame2"
      Tab(1).Control(7)=   "Text3"
      Tab(1).Control(8)=   "chkDrugs"
      Tab(1).Control(9)=   "cmbDoc"
      Tab(1).Control(10)=   "cmdFind"
      Tab(1).ControlCount=   11
      Begin VB.CommandButton cmdFind 
         Height          =   255
         Left            =   -71040
         TabIndex        =   26
         Top             =   480
         Width           =   375
      End
      Begin VB.ComboBox cmbDoc 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmMedReg.frx":087D
         Left            =   -68760
         List            =   "frmMedReg.frx":0887
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   120
         Width           =   3615
      End
      Begin VB.CheckBox chkDrugs 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Issue  prescription"
         Height          =   255
         Left            =   -74760
         TabIndex        =   20
         Top             =   120
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -72840
         TabIndex        =   15
         Top             =   5040
         Width           =   11895
      End
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   -74880
         TabIndex        =   11
         Top             =   5400
         Width           =   13935
         Begin VB.CommandButton cmdConDrugs 
            Caption         =   "Confirm &Drugs"
            Enabled         =   0   'False
            Height          =   375
            Left            =   11880
            TabIndex        =   12
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   120
         TabIndex        =   9
         Top             =   5520
         Width           =   14175
         Begin VB.CheckBox chkDocReq 
            Caption         =   "&Doctor Consultation Required"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   2895
         End
         Begin VB.CommandButton cmdConSymp 
            Caption         =   "&Next >>"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   12120
            TabIndex        =   10
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.TextBox txtVar1 
         Height          =   285
         Left            =   2280
         TabIndex        =   8
         Top             =   120
         Width           =   2055
      End
      Begin MSFlexGridLib.MSFlexGrid MSFSymptoms 
         Height          =   5055
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   8916
         _Version        =   393216
         AllowBigSelection=   0   'False
         HighLight       =   2
         SelectionMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid MSFEmpSymptoms 
         Height          =   5055
         Left            =   4560
         TabIndex        =   6
         Top             =   480
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   8916
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid MSFDrugs 
         Height          =   4095
         Left            =   -74880
         TabIndex        =   13
         Top             =   840
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   7223
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid MSFEmpDrugs 
         Height          =   4095
         Left            =   -70440
         TabIndex        =   14
         Top             =   840
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   7223
         _Version        =   393216
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Recomended Medicines"
         Height          =   255
         Left            =   -70440
         TabIndex        =   25
         Top             =   480
         Width           =   9495
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Medicine Inventory"
         Height          =   255
         Left            =   -74760
         TabIndex        =   24
         Top             =   480
         Width           =   3615
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Consultation Staus"
         Height          =   255
         Left            =   -70320
         TabIndex        =   23
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Special Remarks"
         Height          =   255
         Left            =   -74760
         TabIndex        =   16
         Top             =   5040
         Width           =   1695
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "DIAGNOSIS CARD"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   4560
         TabIndex        =   7
         Top             =   120
         Width           =   9615
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Diagnosis Symptoms"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   180
         Width           =   2055
      End
   End
   Begin MSComctlLib.StatusBar stat1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   8520
      Width           =   15930
      _ExtentX        =   28099
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "10/20/2020"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "5:08 PM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   35278
            MinWidth        =   35278
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtEmpMo 
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label lblFName 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2040
      TabIndex        =   29
      Top             =   1200
      Width           =   12495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Full Name"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Employe Number"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Medical Register"
      BeginProperty Font 
         Name            =   "Neuropolitical Rg"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   14  'Copy Pen
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   2160
      Top             =   0
      Width           =   13815
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   0
      Picture         =   "frmMedReg.frx":08B3
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "frmMedReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Patient_Emp_No As String

Private Sub cmdConSymp_Click()
    If Trim(lblFName.Caption) = "" Then
        MsgBox "Employee Number NOT Entered", vbExclamation
        txtEmpMo.SetFocus
        Exit Sub
    End If

    Dim docreq As Integer
    If chkDocReq.Value = 1 Then
        docreq = 1
    Else
        docreq = 0
    End If
    Set RS = New ADODB.Recordset
    RS.Open "UPDATE HRS_HLC_Emp_Symptoms SET Doctor_Req=" & docreq & ",Symp_Confirmation=1 Where Emp_Number='" & Patient_Emp_No & "' and Inquiry_Number=0", HRS, adOpenStatic, adLockOptimistic
    cmdConDrugs.Enabled = True
    MsgBox "Patient Diagnosis Confirmed", vbInformation
    cmdConSymp.Enabled = False
    PR_Emp_Symptoms
End Sub
Private Sub cmdExit_Click()
    PR_HRS_Close_CON
    Patient_Emp_No = ""
    Close All
    Unload Me
End Sub

Private Sub Form_Load()
    PR_HRS_Open_CON
    PR_Symptoms
    PR_Emp_Symptoms
    PR_Fill_Drugs
    PR_Fill_Emp_Drugs
    stat1.Panels(4) = "Medicle Officer - " + Medi_U_Name
End Sub

Public Sub PR_Symptoms()
    MSFSymptoms.Cols = 2
    Set RS = New ADODB.Recordset
    If Trim(txtVar1.Text) = "" Then
        RS.Open "SELECT * FROM HRS_HLC_MSTR_Symptoms Order by Sym_ID", HRS, adOpenKeyset, adLockReadOnly
    Else
        RS.Open "SELECT * FROM HRS_HLC_MSTR_Symptoms where Sym_Description like '%" & Trim(txtVar1.Text) & "%' Order by Sym_ID", HRS, adOpenKeyset, adLockReadOnly
    End If
    R = 1
    Do While RS.EOF = False
        MSFSymptoms.Rows = R + 1
        MSFSymptoms.TextMatrix(R, 0) = Trim(RS!Sym_ID)
        MSFSymptoms.TextMatrix(R, 1) = Trim(RS!Sym_Description)
        If Val(RS!Sym_Priority) = 1 Then
            MSFSymptoms.Row = R
            MSFSymptoms.Col = 1
            MSFSymptoms.CellBackColor = &HC0C0FF
        End If
        If Val(RS!Sym_Priority) = 2 Then
            MSFSymptoms.Row = R
            MSFSymptoms.Col = 1
            MSFSymptoms.CellBackColor = &HC0FFFF
        End If
        If Val(RS!Sym_Priority) = 3 Then
            MSFSymptoms.Row = R
            MSFSymptoms.Col = 1
            MSFSymptoms.CellBackColor = &HC0FFC0
        End If
        R = R + 1
        RS.MoveNext
    Loop
    
    MSFSymptoms.ColWidth(0) = 800
    MSFSymptoms.ColWidth(1) = 3500
    
    MSFSymptoms.TextMatrix(0, 0) = "ID"
    MSFSymptoms.TextMatrix(0, 1) = "SYMPTOMS"
End Sub

Private Sub MSFEmpSymptoms_DblClick()
    If MsgBox("Do you wish to continue ?", vbQuestion + vbYesNo, "Delete Record") = vbYes Then
        HRS.Execute "Delete from HRS_HLC_Emp_Symptoms where  Emp_Symp_Index=" & Val(MSFEmpSymptoms.TextMatrix(MSFEmpSymptoms.Row, 2)) & "and Symp_Confirmation=0"
        PR_Emp_Symptoms
    End If
End Sub

Private Sub MSFlexGrid1_Click()

End Sub

Private Sub MSFSymptoms_DblClick()
    If Trim(lblFName.Caption) = "" Then
        MsgBox "Employee Number NOT Found", vbExclamation
        Exit Sub
    End If
    Patient_Emp_No = Trim(txtEmpMo.Text)
    Dim Sym_ID As Long
    Set RS = New ADODB.Recordset
    RS.Open "Select * from HRS_HLC_MSTR_Symptoms Where Sym_Description='" & Trim(MSFSymptoms.TextMatrix(MSFSymptoms.Row, 1)) & "'", HRS, adOpenStatic, adLockReadOnly
    If RS.EOF = False Then
        Sym_ID = Val(RS!Sym_ID)
        'If MsgBox("Do you Want to Add Selected Sysmpoms to the Employee ?", vbQuestion + vbYesNo, "Health Care Module") = vbYes Then
            HRS.Execute "INSERT INTO HRS_HLC_Emp_Symptoms(Inquiry_Number,Emp_Number,Sym_ID,U_ID,Trans_Date_Time)" _
                + "Values(0,'" & Patient_Emp_No & "'," & Sym_ID & ",'" & UserID & "','" & Date + Time & "')"
            PR_Emp_Symptoms
        'End If
    End If
End Sub

Private Sub txtEmpMo_Change()
    Set RS = New ADODB.Recordset
    Patient_Emp_No = Trim(txtEmpMo.Text)
    RS.Open "Select * from HRS_sys_HCM_Employee Where Emp_Number='" & Patient_Emp_No & "'", HRS, adOpenStatic, adLockReadOnly
    If RS.EOF = False Then
        lblFName.Caption = Trim(RS!EMP_SurName)
        PR_Emp_Symptoms
    Else
        lblFName.Caption = ""
        Exit Sub
    End If
End Sub

Private Sub txtVar1_KeyPress(KeyAscii As Integer)
    PR_Symptoms
End Sub

Public Sub PR_Emp_Symptoms()
    MSFEmpSymptoms.Cols = 4
    MSFEmpSymptoms.Rows = 1
    MSFEmpSymptoms.TextMatrix(0, 0) = "No."
    MSFEmpSymptoms.TextMatrix(0, 1) = "SYMPTOMS"
    MSFEmpSymptoms.TextMatrix(0, 3) = "OTHER COMMENTS"
    MSFEmpSymptoms.ColWidth(0) = 0
    MSFEmpSymptoms.ColWidth(1) = 3500
    MSFEmpSymptoms.ColWidth(2) = 0
    MSFEmpSymptoms.ColWidth(3) = 5000
    R = 1
    Set GRDRS = New ADODB.Recordset
    GRDRS.Open "SELECT * FROM HRSV_HLC_EMP_Symptoms WHERE EMP_NUMBER='" & Patient_Emp_No & "' AND INQUIRY_NUMBER=0", HRS, adOpenStatic, adLockReadOnly
    Do While GRDRS.EOF = False
        MSFEmpSymptoms.Rows = R + 1
        MSFEmpSymptoms.TextMatrix(R, 0) = R
        MSFEmpSymptoms.TextMatrix(R, 1) = Trim(GRDRS!Sym_Description)
        MSFEmpSymptoms.TextMatrix(R, 2) = Trim(GRDRS!Emp_Symp_Index)
        'MSFEmpSymptoms.TextMatrix(R, 3) = Trim(GRDRS!Remarks)
        If Val(GRDRS!Sym_Priority) = 1 Then
            MSFEmpSymptoms.Row = R
            MSFEmpSymptoms.Col = 1
            MSFEmpSymptoms.CellBackColor = &HC0C0FF
        End If
        If Val(GRDRS!Sym_Priority) = 2 Then
            MSFEmpSymptoms.Row = R
            MSFEmpSymptoms.Col = 1
            MSFEmpSymptoms.CellBackColor = &HC0FFFF
        End If
        If Val(GRDRS!Sym_Priority) = 3 Then
            MSFEmpSymptoms.Row = R
            MSFEmpSymptoms.Col = 1
            MSFEmpSymptoms.CellBackColor = &HC0FFC0
        End If
        R = R + 1
        GRDRS.MoveNext
    Loop
End Sub

Public Sub PR_Fill_Drugs()
    Set GRDRS = New ADODB.Recordset
    GRDRS.Open "Select * from HRS_HLC_MSTR_Medicine Order by Medi_ID", HRS, adOpenStatic, adLockReadOnly
    Do While GRDRS.EOF = False
        MSFDrugs.Rows = R + 1
        MSFDrugs.TextMatrix(R, 0) = Trim(GRDRS!Medi_ID)
        MSFDrugs.TextMatrix(R, 1) = Trim(GRDRS!Medi_Name)
        If GRDRS!Doc_allowed = True Then
            MSFDrugs.Row = R
            MSFDrugs.Col = 1
            MSFDrugs.CellBackColor = &HC0C0FF
        Else
            MSFDrugs.Row = R
            MSFDrugs.Col = 1
            MSFDrugs.CellBackColor = &H80000005
        End If
        R = R + 1
        GRDRS.MoveNext
    Loop
    
    MSFDrugs.ColWidth(0) = 800
    MSFDrugs.ColWidth(1) = 3500
    
    MSFDrugs.TextMatrix(0, 0) = "ID."
    MSFDrugs.TextMatrix(0, 1) = "MEDICINE NAME"
End Sub

Public Sub PR_Fill_Emp_Drugs()
    Set GRDRS = New ADODB.Recordset
    MSFEmpDrugs.Cols = 4
    MSFEmpDrugs.TextMatrix(0, 0) = "NO."
    MSFEmpDrugs.TextMatrix(0, 1) = "MEDICIN NAME"
    MSFEmpDrugs.TextMatrix(0, 2) = "FREQUENCY"
    MSFEmpDrugs.TextMatrix(0, 3) = "DOSE"
    
    MSFEmpDrugs.ColWidth(0) = 800
    MSFEmpDrugs.ColWidth(1) = 3500
    MSFEmpDrugs.ColWidth(2) = 2500
    MSFEmpDrugs.ColWidth(3) = 2500
End Sub
