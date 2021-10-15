VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmMealCsv 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Meal CSV Generation"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6405
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
   ScaleHeight     =   2280
   ScaleWidth      =   6405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Height          =   855
      Left            =   5400
      Picture         =   "frmCsv.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Exit from the System"
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "&Generate"
      Height          =   855
      Left            =   4080
      TabIndex        =   7
      Top             =   1080
      Width           =   1215
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   2025
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
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
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPEndDate 
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   1560
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      Format          =   54394881
      CurrentDate     =   43966
   End
   Begin MSComCtl2.DTPicker DTPStartDate 
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   1080
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      Format          =   54394881
      CurrentDate     =   43966
   End
   Begin MSDataListLib.DataCombo cmbMealType 
      Height          =   315
      Left            =   1320
      TabIndex        =   3
      Top             =   600
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dcmbCluster 
      Height          =   315
      Left            =   1320
      TabIndex        =   10
      Top             =   120
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label4 
      Caption         =   "Cluster Code"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "End Date"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Start Date"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Meal Type"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "frmMealCsv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    Close All
    Unload Me
End Sub

Private Sub cmdGenerate_Click()
On Error GoTo er_EH:
PR_HRS_Open_CON
    Open "C:\Attendance\Meal\Meal_" + Trim(dcmbCluster.Text) + "_" + Format(DTPStartDate.Value, "YYMMdd") + "_" & Trim(cmbMealType.Text) & ".txt" For Output As #1
    Dim E_No As String
    If Trim(cmbMealType.Text) = "" Then
        MsgBox "Meal Type NOT Selected", vbCritical
        Exit Sub
    End If
    
    Set RS = New ADODB.Recordset
    If Trim(cmbMealType.Text) = "Lunch" Then
        RS.Open "SELECT * FROM [dbo].[ADMS_BRDX_Source] Where Cluster_Code='" & Trim(dcmbCluster.Text) & "' and Trans_Date>='" & DTPStartDate.Value & "' and Trans_Date<='" & DTPEndDate.Value & "' and Trans_Time>=Lu_Start_Time and  Trans_Time<=Lu_End_Time", HRS, adOpenKeyset, adLockPessimistic
    End If
    
    If Trim(cmbMealType.Text) = "Snack" Then
        RS.Open "SELECT * FROM [dbo].[ADMS_BRDX_Source] Where Cluster_Code='" & Trim(dcmbCluster.Text) & "'and Trans_Date>='" & DTPStartDate.Value & "' and Trans_Date<='" & DTPEndDate.Value & "' and Trans_Time>=ES_Start_Time and  Trans_Time<=ES_End_Time", HRS, adOpenKeyset, adLockPessimistic
    End If
    
    If Trim(cmbMealType.Text) = "Breakfast" Then
        RS.Open "SELECT * FROM [dbo].[ADMS_BRDX_Source] Where Cluster_Code='" & Trim(dcmbCluster.Text) & "' and Trans_Date>='" & DTPStartDate.Value & "' and Trans_Date<='" & DTPEndDate.Value & "' and Trans_Time>=BF_Start_Time and  Trans_Time<=BF_End_Time", HRS, adOpenKeyset, adLockPessimistic
    End If
    
    Do While RS.EOF = False
        E_No = Format(Trim(RS!Employee_Number), "00000#")
        Print #1, Format(Str(RS!Clock_ID), "00# ") + ":" + E_No + ":" + Format(RS!transaction_Date_Time, "YYMMdd2") + ":" + Format(RS!transaction_Date_Time, "HHmmss")
        RS.MoveNext
    Loop
    Close #1
    
    MsgBox "CSV Generated Successfully", vbInformation
    PR_HRS_Close_CON
    Exit Sub
    
er_EH:
    Close #1
    MsgBox Err.Description, vbCritical
    Err.Clear
End Sub

Private Sub Form_Load()
    PR_HRS_Open_CON
    DTPStartDate.Value = Date
    DTPEndDate.Value = Date
    
    Set RS = New ADODB.Recordset
    RS.Open "Select * from ADMS_MSTR_Meal", HRS, adOpenStatic, adLockReadOnly
    cmbMealType.ListField = "Meal_Type"
    Set cmbMealType.RowSource = RS
    
    Set RS = New ADODB.Recordset
    RS.Open "Select Cluster_ID from ADMS_Clock_Master Group by Cluster_ID", HRS, adOpenStatic, adLockReadOnly
    dcmbCluster.ListField = "Cluster_ID"
    Set dcmbCluster.RowSource = RS
    
End Sub
