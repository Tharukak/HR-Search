VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmEmpBarcode 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Employee Barcode Creator"
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7890
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
   ScaleHeight     =   8385
   ScaleWidth      =   7890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDeleteAll 
      Caption         =   "&Delete All"
      Height          =   255
      Left            =   6360
      TabIndex        =   6
      Top             =   1080
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid MSFEmpDetails 
      Height          =   6135
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "Double click on the Employee Number to Delete"
      Top             =   1440
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   10821
      _Version        =   393216
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtBarcodeNo 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
   End
   Begin Crystal.CrystalReport CRW 
      Left            =   600
      Top             =   600
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
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "&Generate"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   7680
      Width           =   7695
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   8130
      Width           =   7890
      _ExtentX        =   13917
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "12/13/2020"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "12:30 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Enabled         =   0   'False
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Barcode Number"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   -2520
      Picture         =   "frmEmpBarcode.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10455
   End
End
Attribute VB_Name = "frmEmpBarcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Emp_BarcodeNo As String
Private Sub cmdAdd_Click()
On Error GoTo er_EH:
    Dim EPF_No As String
    Dim EMP_FName As String
    Dim EMP_SurName As String
    Dim EMP_NIC_No As String
    Dim Dsg_Name As String
    Dim Mechine_ID As String
    
    If Trim(txtBarcodeNo.Text) = "" Then
        MsgBox "Employee Barcode Number NOT Entered", vbInformation
        Exit Sub
    Else
        PR_HRS_Open_CON
        Emp_BarcodeNo = Trim(txtBarcodeNo.Text)
        Set RS = New ADODB.Recordset
        RS.Open "Select * from HCM_sys_MSTR_Employee Where Emp_Barcodeno='" & Emp_BarcodeNo & "'", con_ADMS, adOpenStatic, adLockReadOnly
        If RS.EOF = False Then
            EPF_No = Trim(RS!Epf_Number)
            EMP_FName = Trim(RS!Emp_FirstName)
            EMP_SurName = Trim(RS!EMP_SurName)
            Dsg_Name = Trim(RS!Dsg_Name)
            EMP_NIC_No = Trim(RS!EMP_NIC_No)
            HRS.Execute "INSERT INTO HRS_HR_Emp_Barcode(Emp_BarcodeNo,EPF_Number,Emp_FirstName,Emp_SurName,Emp_NIC_No,Dsg_Name,Mechine_ID) VALUES('" & Emp_BarcodeNo & "','" & EPF_No & "','" & EMP_FName & "','" & EMP_SurName & "','" & EMP_NIC_No & "','" & Dsg_Name & "','" & VBA.Environ("COMPUTERNAME") & "')"
            txtBarcodeNo.Text = ""
            PR_FMT_Grid
        Else
            MsgBox "Invalid Employee Barcode Number", vbExclamation
            txtBarcodeNo.Text = ""
        End If
    End If
    
    Exit Sub
er_EH:
    MsgBox Err.Description, vbCritical
    Err.Clear
End Sub

Private Sub cmdDeleteAll_Click()
On Error GoTo er_EH:
    If MsgBox("Are you sure Do you Want to Delete All Records ?", vbQuestion + vbYesNo, "Delete Record") = vbYes Then
        HRS.Execute "Delete from HRS_HR_Emp_Barcode Where Mechine_ID='" & VBA.Environ("COMPUTERNAME") & "'"
        PR_FMT_Grid
    End If
    Exit Sub
    
er_EH:
    MsgBox Err.Description, vbCritical
    Err.Clear
End Sub

Private Sub cmdGenerate_Click()
On Error GoTo er_EH:
    PR_REPORT_PATH
    Dim Report_Name As String
    PR_HRS_Open_CON
    PR_REPORT_PATH
    CRW.ReportFileName = Report_Path + "HRS_E_RPT0012.rpt"
    CRW.ParameterFields(0) = "Mechine_ID;" & VBA.Environ("COMPUTERNAME") & "; true"
    CRW.Connect = DSN_SETTINGS
    CRW.Action = 1
    Exit Sub
    
er_EH:
    MsgBox Err.Description, vbCritical
    Err.Clear
End Sub

Private Sub Form_Load()
    ADMS_Connect
    PR_FMT_Grid
End Sub

Public Sub PR_FMT_Grid()
    Set GRDRS = New ADODB.Recordset
    PR_HRS_Open_CON
    GRDRS.Open "Select * from HRS_HR_Emp_Barcode Where Mechine_ID='" & VBA.Environ("COMPUTERNAME") & "'", HRS, adOpenStatic, adLockReadOnly
    MSFEmpDetails.Cols = 6
    MSFEmpDetails.Rows = 1
    R = 1
    Do While GRDRS.EOF = False
        MSFEmpDetails.Rows = Val(GRDRS.RecordCount) + 1
        MSFEmpDetails.TextMatrix(R, 0) = Trim(GRDRS!Emp_BarcodeNo)
        MSFEmpDetails.TextMatrix(R, 1) = Trim(GRDRS!Epf_Number)
        MSFEmpDetails.TextMatrix(R, 2) = Trim(GRDRS!Emp_FirstName)
        MSFEmpDetails.TextMatrix(R, 3) = Trim(GRDRS!EMP_SurName)
        MSFEmpDetails.TextMatrix(R, 4) = Trim(GRDRS!EMP_NIC_No)
        MSFEmpDetails.TextMatrix(R, 5) = Trim(GRDRS!Dsg_Name)
        R = R + 1
        GRDRS.MoveNext
    Loop
    
    MSFEmpDetails.ColWidth(0) = 800
    MSFEmpDetails.ColWidth(1) = 800
    MSFEmpDetails.ColWidth(2) = 2000
    MSFEmpDetails.ColWidth(3) = 2000
    MSFEmpDetails.ColWidth(4) = 1000
    MSFEmpDetails.ColWidth(5) = 2000
    
    MSFEmpDetails.TextMatrix(0, 0) = "Barcode#"
    MSFEmpDetails.TextMatrix(0, 1) = "EPF#"
    MSFEmpDetails.TextMatrix(0, 2) = "First Name"
    MSFEmpDetails.TextMatrix(0, 3) = "Sur Name"
    MSFEmpDetails.TextMatrix(0, 4) = "NIC"
    MSFEmpDetails.TextMatrix(0, 5) = "Designation"
    
End Sub
Private Sub MSFEmpDetails_DblClick()
    On Error GoTo er_EH:
    If MsgBox("Are you sure Do you Want to Delete All Records ?", vbQuestion + vbYesNo, "Delete Record") = vbYes Then
        HRS.Execute "Delete from HRS_HR_Emp_Barcode Where Mechine_ID='" & VBA.Environ("COMPUTERNAME") & "' and Emp_Barcodeno='" & MSFEmpDetails.TextMatrix(MSFEmpDetails.Row, 0) & "'"
        PR_FMT_Grid
    End If
    Exit Sub
    
er_EH:
    MsgBox Err.Description, vbCritical
    Err.Clear
End Sub

Private Sub txtBarcodeNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdAdd_Click
    End If
End Sub
