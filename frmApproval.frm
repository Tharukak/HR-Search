VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmApproval 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Approval Pannel"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14850
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
   ScaleHeight     =   8580
   ScaleWidth      =   14850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   7335
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   12938
      _Version        =   393216
      TabsPerRow      =   5
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&VR Approval"
      TabPicture(0)   =   "frmApproval.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "MSFApproval"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Escalated Approval"
      TabPicture(1)   =   "frmApproval.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "MSFEscalation"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "&Special Approval"
      TabPicture(2)   =   "frmApproval.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "MSFSPApprove"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame4"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   -74880
         TabIndex        =   26
         Top             =   6480
         Width           =   13575
         Begin VB.CommandButton cmdSPApprovedAll 
            Caption         =   "&Approve All"
            Height          =   375
            Left            =   10320
            TabIndex        =   37
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdSPDisApproved 
            Caption         =   "&Dis Approved"
            Height          =   375
            Left            =   12120
            TabIndex        =   30
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmdSPApproved 
            Caption         =   "&Approve"
            Height          =   375
            Left            =   9000
            TabIndex        =   29
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdSP 
            Caption         =   "&Refresh"
            Height          =   375
            Left            =   3240
            TabIndex        =   28
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1560
            TabIndex        =   27
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label12 
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   7560
            TabIndex        =   34
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label11 
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   6240
            TabIndex        =   33
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label10 
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   4920
            TabIndex        =   32
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label9 
            Caption         =   "Transaction Year"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   -74880
         TabIndex        =   16
         Top             =   6480
         Width           =   13575
         Begin VB.CommandButton cmdEsApprovedAll 
            Caption         =   "&Approve All"
            Height          =   375
            Left            =   10320
            TabIndex        =   36
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1560
            TabIndex        =   20
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton cmdEsc 
            Caption         =   "&Refresh"
            Height          =   375
            Left            =   3240
            TabIndex        =   19
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton cmdEsApproved 
            Caption         =   "&Approve"
            Height          =   375
            Left            =   9000
            TabIndex        =   18
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdVxDisApproved 
            Caption         =   "&Dis Approved"
            Height          =   375
            Left            =   12120
            TabIndex        =   17
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label8 
            Caption         =   "Transaction Year"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label7 
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   4920
            TabIndex        =   23
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label6 
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   6240
            TabIndex        =   22
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label5 
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   7560
            TabIndex        =   21
            Top             =   240
            Width           =   1335
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSFEscalation 
         Height          =   6015
         Left            =   -74880
         TabIndex        =   15
         Top             =   480
         Width           =   13575
         _ExtentX        =   23945
         _ExtentY        =   10610
         _Version        =   393216
         BackColor       =   14286847
         ForeColor       =   64
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   120
         TabIndex        =   5
         Top             =   6480
         Width           =   13575
         Begin VB.CommandButton cmdApproveAll 
            Caption         =   "&Approve All"
            Height          =   375
            Left            =   10320
            TabIndex        =   35
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdDisApprove 
            Caption         =   "&Dis Approved"
            Height          =   375
            Left            =   12120
            TabIndex        =   13
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmdApprove 
            Caption         =   "&Approve"
            Height          =   375
            Left            =   9000
            TabIndex        =   12
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "&Refresh"
            Height          =   375
            Left            =   3240
            TabIndex        =   8
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox txtTransYear 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1560
            TabIndex        =   7
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label4 
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   7560
            TabIndex        =   11
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label3 
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   6240
            TabIndex        =   10
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label2 
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   4920
            TabIndex        =   9
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Transaction Year"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   1455
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSFApproval 
         Height          =   6015
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   13575
         _ExtentX        =   23945
         _ExtentY        =   10610
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid MSFSPApprove 
         Height          =   6015
         Left            =   -74880
         TabIndex        =   25
         Top             =   480
         Width           =   13575
         _ExtentX        =   23945
         _ExtentY        =   10610
         _Version        =   393216
         BackColor       =   14286847
         ForeColor       =   64
      End
   End
   Begin VB.Frame Frame2 
      Height          =   7335
      Left            =   14040
      TabIndex        =   1
      Top             =   840
      Width           =   735
      Begin VB.CommandButton cmdExit 
         Height          =   495
         Left            =   120
         Picture         =   "frmApproval.frx":0054
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Exit from the System"
         Top             =   6720
         Width           =   495
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   8280
      Width           =   14850
      _ExtentX        =   26194
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   11
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
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel11 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Approval Pannel"
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
      Left            =   2400
      TabIndex        =   14
      Top             =   120
      Width           =   5655
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   14  'Copy Pen
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   2280
      Top             =   0
      Width           =   12615
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   0
      Picture         =   "frmApproval.frx":0899
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "frmApproval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Req_ID As Long
Dim ExReq_ID As Long
Private Sub cmdApprove_Click()
    Module_ID = 1
    Sub_Module_ID = 1
    
    If Req_ID = 0 Then
        MsgBox "VR Number NOT Selected", vbExclamation
        Exit Sub
    Else
        PR_HRS_Open_CON
        Set RS = New ADODB.Recordset
        RS.Open "Select * from HRSV_TR_Requests Where VR_No=" & Req_ID, HRS, adOpenStatic, adLockReadOnly
        If RS.EOF = False Then
            FN_Status_Update 2, Req_ID, Module_ID, Sub_Module_ID
            PR_FMT_Grid
        Else
            MsgBox "VR Number NOT Found", vbExclamation
            Exit Sub
        End If
        PR_HRS_Close_CON
    End If
End Sub

Private Sub cmdApproveAll_Click()
    Module_ID = 1
    Sub_Module_ID = 1
    
    For R = 1 To MSFApproval.Rows - 1
        If Req_ID = 0 Then
            MsgBox "VR Number NOT Selected", vbExclamation
            Exit Sub
        Else
            Req_ID = MSFApproval.TextMatrix(R, 0)
            PR_HRS_Open_CON
            Set RS = New ADODB.Recordset
            RS.Open "Select * from HRSV_TR_Requests Where VR_No=" & Req_ID, HRS, adOpenStatic, adLockReadOnly
            If RS.EOF = False Then
                FN_Status_Update 2, Req_ID, Module_ID, Sub_Module_ID
                PR_FMT_Grid
            Else
                MsgBox "VR Number NOT Found", vbExclamation
                Exit Sub
            End If
            PR_HRS_Close_CON
        End If
    Next R
End Sub

Private Sub cmdDisApprove_Click()
    Module_ID = 1
    Sub_Module_ID = 1
    If Req_ID = 0 Then
        MsgBox "VR Number NOT Selected", vbExclamation
        Exit Sub
    Else
        PR_HRS_Open_CON
        Set RS = New ADODB.Recordset
        RS.Open "Select * from HRSV_TR_Requests Where VR_No=" & Req_ID, HRS, adOpenStatic, adLockReadOnly
        If RS.EOF = False Then
            FN_Status_Update 3, Req_ID, Module_ID, Sub_Module_ID
            PR_FMT_Grid
        Else
            MsgBox "VR Number NOT Found", vbExclamation
            Exit Sub
        End If
        PR_HRS_Close_CON
    End If
End Sub

Private Sub cmdEsApproved_Click()
    Module_ID = 1
    Sub_Module_ID = 1
    If ExReq_ID = 0 Then
        MsgBox "VR Number NOT Selected", vbExclamation
        Exit Sub
    Else
        PR_HRS_Open_CON
        Set RS = New ADODB.Recordset
        RS.Open "Select * from HRSV_TR_Requests Where VR_No=" & ExReq_ID, HRS, adOpenStatic, adLockReadOnly
        If RS.EOF = False Then
            FN_Status_Update 3, ExReq_ID, Module_ID, Sub_Module_ID
            PR_FMT_Escalation_Grid
        Else
            MsgBox "VR Number NOT Found", vbExclamation
            Exit Sub
        End If
        PR_HRS_Close_CON
    End If
End Sub

Private Sub cmdEsApprovedAll_Click()
    Module_ID = 1
    Sub_Module_ID = 1
    
    For R = 1 To MSFEscalation.Rows - 1
        Req_ID = MSFEscalation.TextMatrix(R, 0)
        PR_HRS_Open_CON
        Set RS = New ADODB.Recordset
        RS.Open "Select * from HRSV_TR_Requests Where VR_No=" & Req_ID, HRS, adOpenStatic, adLockReadOnly
        If RS.EOF = False Then
            FN_Status_Update 3, ExReq_ID, Module_ID, Sub_Module_ID
            PR_FMT_Escalation_Grid
        Else
            MsgBox "VR Number NOT Found", vbExclamation
            Exit Sub
        End If
        PR_HRS_Close_CON
    Next R
End Sub

Private Sub cmdEsc_Click()
    If Trim(txtTransYear.Text) = "" Then
        MsgBox "Transaction Year NOT Entered", vbExclamation
        Exit Sub
    Else
        PR_FMT_Escalation_Grid
    End If
End Sub

Private Sub cmdExit_Click()
    Close All
    Unload Me
End Sub

Private Sub cmdRefresh_Click()
    If Trim(txtTransYear.Text) = "" Then
        MsgBox "Transaction Year NOT Entered", vbExclamation
        Exit Sub
    Else
        Call PR_FMT_Grid
    End If
End Sub

Private Sub cmdSP_Click()
    If Trim(txtTransYear.Text) = "" Then
        MsgBox "Transaction Year NOT Entered", vbExclamation
        Exit Sub
    Else
        Call PR_FMT_SPApprove_Grid
    End If
End Sub

Private Sub cmdSPApproved_Click()
    Module_ID = 1
    Sub_Module_ID = 1
    
    If Req_ID = 0 Then
        MsgBox "VR Number NOT Selected", vbExclamation
        Exit Sub
    Else
        PR_HRS_Open_CON
        Set RS = New ADODB.Recordset
        RS.Open "Select * from HRSV_TR_Requests Where VR_No=" & Req_ID, HRS, adOpenStatic, adLockReadOnly
        If RS.EOF = False Then
            FN_Status_Update 2, Req_ID, Module_ID, Sub_Module_ID
            PR_FMT_SPApprove_Grid
        Else
            MsgBox "VR Number NOT Found", vbExclamation
            Exit Sub
        End If
        PR_HRS_Close_CON
    End If
End Sub

Private Sub cmdSPApprovedAll_Click()
    Module_ID = 1
    Sub_Module_ID = 1
    
    For R = 1 To MSFSPApprove.Rows - 1
        If Req_ID = 0 Then
            MsgBox "VR Number NOT Selected", vbExclamation
            Exit Sub
        Else
            Req_ID = MSFSPApprove.TextMatrix(R, 0)
            PR_HRS_Open_CON
            Set RS = New ADODB.Recordset
            RS.Open "Select * from HRSV_TR_Requests Where VR_No=" & Req_ID, HRS, adOpenStatic, adLockReadOnly
            If RS.EOF = False Then
                FN_Status_Update 3, Req_ID, Module_ID, Sub_Module_ID
                PR_FMT_SPApprove_Grid
            Else
                MsgBox "VR Number NOT Found", vbExclamation
                Exit Sub
            End If
            PR_HRS_Close_CON
        End If
    Next R
End Sub

Private Sub cmdVxDisApproved_Click()
    Module_ID = 1
    Sub_Module_ID = 1
    If ExReq_ID = 0 Then
        MsgBox "VR Number NOT Selected", vbExclamation
        Exit Sub
    Else
        PR_HRS_Open_CON
        Set RS = New ADODB.Recordset
        RS.Open "Select * from HRSV_TR_Requests Where VR_No=" & ExReq_ID, HRS, adOpenStatic, adLockReadOnly
        If RS.EOF = False Then
            'FN_Status_Update 3, Ex_Req_ID, Module_ID, Sub_Module_ID
            PR_FMT_Escalation_Grid
        Else
            MsgBox "VR Number NOT Found", vbExclamation
            Exit Sub
        End If
        PR_HRS_Close_CON
    End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
    txtTransYear.Text = Format(Date, "yyyy")
    PR_FMT_Grid
    PR_FMT_Escalation_Grid
End Sub
Public Sub PR_FMT_Grid()
    PR_HRS_Open_CON
        MSFApproval.Cols = 10
        Set GRDRS = New ADODB.Recordset
        R = 1
        MSFApproval.Clear
        GRDRS.Open "Select * from HRSV_TR_Requests Where R_ID='" & UserID & "' and year(T_Date_Time)=" & Val(txtTransYear.Text) & "and Status_ID=1", HRS, adOpenStatic, adLockReadOnly
        If GRDRS.EOF = False Then
            Do While GRDRS.EOF = False
                MSFApproval.Rows = R + 1
                MSFApproval.TextMatrix(R, 0) = Val(GRDRS!VR_No)
                MSFApproval.TextMatrix(R, 1) = Format(GRDRS!Req_Date_Time, "dd-MMM-yyyy")
                MSFApproval.TextMatrix(R, 2) = Format(GRDRS!Req_Date_Time, "HH:mm:ss")
                MSFApproval.TextMatrix(R, 3) = Val(GRDRS!Distance)
                MSFApproval.TextMatrix(R, 4) = Format(Val(GRDRS!Distance) * Val(GRDRS!AC_Rate), "###,###.00")
                MSFApproval.TextMatrix(R, 5) = Trim(GRDRS!Req_ComCode)
                MSFApproval.TextMatrix(R, 6) = Trim(GRDRS!To_City)
                MSFApproval.TextMatrix(R, 7) = Trim(GRDRS!To_Loc)
                MSFApproval.TextMatrix(R, 8) = Trim(GRDRS!Loc_Dtls)
                MSFApproval.TextMatrix(R, 9) = Trim(GRDRS!U_DName)
                R = R + 1
                GRDRS.MoveNext
            Loop
        Else
            MSFApproval.Rows = 1
        End If
        
        MSFApproval.ColWidth(0) = 800
        MSFApproval.ColWidth(1) = 1000
        MSFApproval.ColWidth(2) = 1000
        MSFApproval.ColWidth(3) = 1000
        MSFApproval.ColWidth(4) = 800
        MSFApproval.ColWidth(5) = 800
        MSFApproval.ColWidth(6) = 2000
        MSFApproval.ColWidth(7) = 2000
        MSFApproval.ColWidth(8) = 2000
        MSFApproval.ColWidth(9) = 2000
        
        MSFApproval.TextMatrix(0, 0) = "VR No."
        MSFApproval.TextMatrix(0, 1) = "Req. Date"
        MSFApproval.TextMatrix(0, 2) = "Req. Time"
        MSFApproval.TextMatrix(0, 3) = "Km"
        MSFApproval.TextMatrix(0, 4) = "Rs."
        MSFApproval.TextMatrix(0, 5) = "From"
        MSFApproval.TextMatrix(0, 6) = "To City"
        MSFApproval.TextMatrix(0, 7) = "To Location"
        MSFApproval.TextMatrix(0, 8) = "Location Details"
        MSFApproval.TextMatrix(0, 9) = "Requester"
    PR_HRS_Close_CON
End Sub

Private Sub MSFApproval_Click()
        Req_ID = Val(MSFApproval.TextMatrix(MSFApproval.Row, 0))
End Sub

Public Sub PR_FMT_Escalation_Grid()
    MSFEscalation.Cols = 10
    MSFEscalation.Rows = 1
    R = 1
    Set RS = New ADODB.Recordset
    PR_HRS_Open_CON
    RS.Open "Select * from HRS_sys_MSTR_Employee Where U_ID='" & UserID & "' and SP_Approve=1", HRS, adOpenStatic, adLockReadOnly
    If RS.EOF = False Then
        Set GRDRS = New ADODB.Recordset
        MSFEscalation.Clear
        GRDRS.Open "Select * from HRSV_TR_Requests Where year(T_Date_Time)=" & Val(txtTransYear.Text) & "and Status_ID=7", HRS, adOpenStatic, adLockReadOnly
        If GRDRS.EOF = False Then
            Do While GRDRS.EOF = False
                MSFEscalation.Rows = R + 1
                MSFEscalation.TextMatrix(R, 0) = Val(GRDRS!VR_No)
                MSFEscalation.TextMatrix(R, 1) = Format(GRDRS!Req_Date_Time, "dd-MMM-yyyy")
                MSFEscalation.TextMatrix(R, 2) = Format(GRDRS!Req_Date_Time, "HH:mm:ss")
                MSFEscalation.TextMatrix(R, 3) = Val(GRDRS!Distance)
                MSFEscalation.TextMatrix(R, 4) = Format(Val(GRDRS!Distance) * Val(GRDRS!AC_Rate), "###,###.00")
                MSFEscalation.TextMatrix(R, 5) = Trim(GRDRS!Req_ComCode)
                MSFEscalation.TextMatrix(R, 6) = Trim(GRDRS!To_City)
                MSFEscalation.TextMatrix(R, 7) = Trim(GRDRS!To_Loc)
                MSFEscalation.TextMatrix(R, 8) = Trim(GRDRS!Loc_Dtls)
                MSFEscalation.TextMatrix(R, 9) = Trim(GRDRS!U_DName)
                R = R + 1
                GRDRS.MoveNext
            Loop
        Else
            MSFEscalation.Rows = 1
        End If
    End If
    PR_HRS_Close_CON
    MSFEscalation.ColWidth(0) = 800
    MSFEscalation.ColWidth(1) = 1000
    MSFEscalation.ColWidth(2) = 1000
    MSFEscalation.ColWidth(3) = 1000
    MSFEscalation.ColWidth(4) = 800
    MSFEscalation.ColWidth(5) = 800
    MSFEscalation.ColWidth(6) = 2000
    MSFEscalation.ColWidth(7) = 2000
    MSFEscalation.ColWidth(8) = 2000
    MSFEscalation.ColWidth(9) = 2000
    
    MSFEscalation.TextMatrix(0, 0) = "VR No."
    MSFEscalation.TextMatrix(0, 1) = "Req. Date"
    MSFEscalation.TextMatrix(0, 2) = "Req. Time"
    MSFEscalation.TextMatrix(0, 3) = "Km"
    MSFEscalation.TextMatrix(0, 4) = "Rs."
    MSFEscalation.TextMatrix(0, 5) = "From"
    MSFEscalation.TextMatrix(0, 6) = "To City"
    MSFEscalation.TextMatrix(0, 7) = "To Location"
    MSFEscalation.TextMatrix(0, 8) = "Location Details"
    MSFEscalation.TextMatrix(0, 9) = "Requester"
End Sub

Private Sub MSFEscalation_Click()
    ExReq_ID = Val(MSFEscalation.TextMatrix(MSFEscalation.Row, 0))
End Sub

Public Sub PR_FMT_SPApprove_Grid()
    MSFSPApprove.Cols = 10
    MSFSPApprove.Rows = 1
    R = 1
    Set RS = New ADODB.Recordset
    PR_HRS_Open_CON
    RS.Open "Select * from HRS_sys_MSTR_Employee Where U_ID='" & UserID & "' and SP_Approve=1", HRS, adOpenStatic, adLockReadOnly
    If RS.EOF = False Then
        Set GRDRS = New ADODB.Recordset
        MSFSPApprove.Clear
        GRDRS.Open "Select * from HRSV_TR_Requests Where year(T_Date_Time)=" & Val(txtTransYear.Text) & "and Status_ID=8", HRS, adOpenStatic, adLockReadOnly
        If GRDRS.EOF = False Then
            Do While GRDRS.EOF = False
                MSFSPApprove.Rows = R + 1
                MSFSPApprove.TextMatrix(R, 0) = Val(GRDRS!VR_No)
                MSFSPApprove.TextMatrix(R, 1) = Format(GRDRS!Req_Date_Time, "dd-MMM-yyyy")
                MSFSPApprove.TextMatrix(R, 2) = Format(GRDRS!Req_Date_Time, "HH:mm:ss")
                MSFSPApprove.TextMatrix(R, 3) = Val(GRDRS!Distance)
                MSFSPApprove.TextMatrix(R, 4) = Format(Val(GRDRS!Distance) * Val(GRDRS!AC_Rate), "###,###.00")
                MSFSPApprove.TextMatrix(R, 5) = Trim(GRDRS!Req_ComCode)
                MSFSPApprove.TextMatrix(R, 6) = Trim(GRDRS!To_City)
                MSFSPApprove.TextMatrix(R, 7) = Trim(GRDRS!To_Loc)
                MSFSPApprove.TextMatrix(R, 8) = Trim(GRDRS!Loc_Dtls)
                MSFSPApprove.TextMatrix(R, 9) = Trim(GRDRS!U_DName)
                R = R + 1
                GRDRS.MoveNext
            Loop
        Else
            MSFSPApprove.Rows = 1
        End If
    End If
    PR_HRS_Close_CON
    MSFSPApprove.ColWidth(0) = 800
    MSFSPApprove.ColWidth(1) = 1000
    MSFSPApprove.ColWidth(2) = 1000
    MSFSPApprove.ColWidth(3) = 1000
    MSFSPApprove.ColWidth(4) = 800
    MSFSPApprove.ColWidth(5) = 800
    MSFSPApprove.ColWidth(6) = 2000
    MSFSPApprove.ColWidth(7) = 2000
    MSFSPApprove.ColWidth(8) = 2000
    MSFSPApprove.ColWidth(9) = 2000
    
    MSFSPApprove.TextMatrix(0, 0) = "VR No."
    MSFSPApprove.TextMatrix(0, 1) = "Req. Date"
    MSFSPApprove.TextMatrix(0, 2) = "Req. Time"
    MSFSPApprove.TextMatrix(0, 3) = "Km"
    MSFSPApprove.TextMatrix(0, 4) = "Rs."
    MSFSPApprove.TextMatrix(0, 5) = "From"
    MSFSPApprove.TextMatrix(0, 6) = "To City"
    MSFSPApprove.TextMatrix(0, 7) = "To Location"
    MSFSPApprove.TextMatrix(0, 8) = "Location Details"
    MSFSPApprove.TextMatrix(0, 9) = "Requester"
End Sub

Private Sub MSFSPApprove_Click()
    Req_ID = Val(MSFSPApprove.TextMatrix(MSFSPApprove.Row, 0))
End Sub
