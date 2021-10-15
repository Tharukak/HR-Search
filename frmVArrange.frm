VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmVArrange 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Vehicle Allocation"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15570
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   15570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport CRW 
      Left            =   6120
      Top             =   840
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
   Begin VB.Frame Frame1 
      Height          =   7455
      Left            =   14400
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove Join"
         Height          =   855
         Left            =   120
         TabIndex        =   18
         Top             =   2760
         Width           =   855
      End
      Begin VB.CommandButton cmbRemove 
         Caption         =   "&Remove Join"
         Height          =   855
         Left            =   -2280
         TabIndex        =   9
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmbJoin 
         Caption         =   "&Join"
         Height          =   855
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   855
      End
      Begin VB.CommandButton cmdExit 
         Height          =   735
         Left            =   120
         Picture         =   "frmVArrange.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Exit from the System"
         Top             =   6600
         Width           =   855
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2415
      Left            =   120
      TabIndex        =   5
      Top             =   6240
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   4260
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   2
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "&Vehicle Details"
      TabPicture(0)   =   "frmVArrange.frx":0845
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblDriver"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblDno1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblDno2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label5"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label9"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblLic"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblIns"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblGrade"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "DcmbStdReason"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "DcmbVno"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmdAllocate"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Frame2"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Frame3"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cmdView"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "&Change Allocation"
      TabPicture(1)   =   "frmVArrange.frx":0861
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture1"
      Tab(1).Control(1)=   "dCmbCVNo"
      Tab(1).Control(2)=   "cmdFind"
      Tab(1).Control(3)=   "txtcAllNo"
      Tab(1).Control(4)=   "lblCNVModel"
      Tab(1).Control(5)=   "lblCNVCat"
      Tab(1).Control(6)=   "lblCNDCont"
      Tab(1).Control(7)=   "lblCNDName"
      Tab(1).Control(8)=   "Label12"
      Tab(1).Control(9)=   "lblCEVModel"
      Tab(1).Control(10)=   "lblCEVCat"
      Tab(1).Control(11)=   "lblCEDCont"
      Tab(1).Control(12)=   "lblCEDName"
      Tab(1).Control(13)=   "Label6"
      Tab(1).ControlCount=   14
      Begin VB.PictureBox Picture1 
         Height          =   1815
         Left            =   -63960
         ScaleHeight     =   1755
         ScaleWidth      =   1995
         TabIndex        =   52
         Top             =   120
         Width           =   2055
         Begin VB.CommandButton cmdChange 
            Caption         =   "&Change Vehicle"
            Height          =   1575
            Left            =   120
            Picture         =   "frmVArrange.frx":087D
            Style           =   1  'Graphical
            TabIndex        =   53
            Top             =   120
            Width           =   1815
         End
      End
      Begin MSDataListLib.DataCombo dCmbCVNo 
         Height          =   330
         Left            =   -67800
         TabIndex        =   47
         Top             =   120
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   582
         _Version        =   393216
         Style           =   2
         Text            =   "DataCombo1"
      End
      Begin VB.CommandButton cmdFind 
         Height          =   255
         Left            =   -71160
         TabIndex        =   45
         Top             =   120
         Width           =   975
      End
      Begin VB.TextBox txtcAllNo 
         Height          =   315
         Left            =   -72960
         TabIndex        =   40
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "&View"
         Height          =   255
         Left            =   4320
         TabIndex        =   37
         Top             =   120
         Width           =   975
      End
      Begin VB.Frame Frame3 
         Caption         =   "Printing Options"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   860
         Left            =   8760
         TabIndex        =   28
         Top             =   1150
         Width           =   5295
         Begin VB.OptionButton Opt1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Printer"
            Height          =   255
            Left            =   3000
            TabIndex        =   30
            Top             =   360
            Width           =   2055
         End
         Begin VB.OptionButton Opt2 
            BackColor       =   &H00E0E0E0&
            Caption         =   " Window"
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   360
            Width           =   2535
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Allocation Inquery"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   860
         Left            =   240
         TabIndex        =   27
         Top             =   1150
         Width           =   8415
         Begin VB.TextBox Text3 
            Height          =   255
            Left            =   4440
            TabIndex        =   36
            Top             =   480
            Width           =   1815
         End
         Begin VB.CommandButton Command2 
            Caption         =   "&Inquire"
            Height          =   495
            Left            =   6360
            TabIndex        =   34
            Top             =   240
            Width           =   1815
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "&Generate"
            Height          =   495
            Left            =   2160
            TabIndex        =   33
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txtAllNo 
            Height          =   255
            Left            =   240
            TabIndex        =   32
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Caption         =   "Requisition Number"
            Height          =   255
            Left            =   4440
            TabIndex        =   35
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Caption         =   "Allocation Number"
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.CommandButton cmdAllocate 
         Caption         =   "&Allocate"
         Height          =   975
         Left            =   12720
         TabIndex        =   26
         Top             =   120
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   1920
         TabIndex        =   22
         Top             =   840
         Width           =   5055
      End
      Begin MSDataListLib.DataCombo DcmbVno 
         Height          =   330
         Left            =   1920
         TabIndex        =   12
         Top             =   120
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   582
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DcmbStdReason 
         Height          =   330
         Left            =   1920
         TabIndex        =   20
         Top             =   480
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   582
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label lblCNVModel 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   -69600
         TabIndex        =   51
         Top             =   1560
         Width           =   4575
      End
      Begin VB.Label lblCNVCat 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   -69600
         TabIndex        =   50
         Top             =   1200
         Width           =   4575
      End
      Begin VB.Label lblCNDCont 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   -69600
         TabIndex        =   49
         Top             =   840
         Width           =   4575
      End
      Begin VB.Label lblCNDName 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   -69600
         TabIndex        =   48
         Top             =   480
         Width           =   4575
      End
      Begin VB.Label Label12 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Vehicle Number"
         Height          =   255
         Left            =   -69600
         TabIndex        =   46
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label lblCEVModel 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000040&
         Height          =   255
         Left            =   -74760
         TabIndex        =   44
         Top             =   1560
         Width           =   4575
      End
      Begin VB.Label lblCEVCat 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000040&
         Height          =   255
         Left            =   -74760
         TabIndex        =   43
         Top             =   1200
         Width           =   4575
      End
      Begin VB.Label lblCEDCont 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000040&
         Height          =   255
         Left            =   -74760
         TabIndex        =   42
         Top             =   840
         Width           =   4575
      End
      Begin VB.Label lblCEDName 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000040&
         Height          =   255
         Left            =   -74760
         TabIndex        =   41
         Top             =   480
         Width           =   4575
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Allocation Number"
         Height          =   255
         Left            =   -74760
         TabIndex        =   39
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label lblGrade 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Grade"
         Height          =   255
         Left            =   10800
         TabIndex        =   25
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblIns 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ins. Status"
         Height          =   255
         Left            =   9000
         TabIndex        =   24
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblLic 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Licence Status"
         Height          =   255
         Left            =   7080
         TabIndex        =   23
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Remarks"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Req. Category"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label lblDno2 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   9720
         TabIndex        =   17
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label lblDno1 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   7080
         TabIndex        =   16
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label lblDriver 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   7080
         TabIndex        =   15
         Top             =   120
         Width           =   5535
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Driver Contact No"
         Height          =   255
         Left            =   5400
         TabIndex        =   14
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "DriverName"
         Height          =   255
         Left            =   5400
         TabIndex        =   13
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Vehicle Number"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   120
         Width           =   1575
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   8700
      Width           =   15570
      _ExtentX        =   27464
      _ExtentY        =   661
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
   Begin VB.CommandButton cmdRefresh 
      Height          =   400
      Left            =   4680
      Picture         =   "frmVArrange.frx":1AE8
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid MSFRequest 
      Height          =   2535
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   4471
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker DTPTDate 
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   840
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      Format          =   101253121
      CurrentDate     =   43307
   End
   Begin MSFlexGridLib.MSFlexGrid MSFAllocated 
      Height          =   2175
      Left            =   120
      TabIndex        =   10
      Top             =   3960
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   3836
      _Version        =   393216
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle Arrangement"
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
      TabIndex        =   38
      Top             =   120
      Width           =   5655
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Vehicle Requested Date"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   0
      Picture         =   "frmVArrange.frx":238E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   14  'Copy Pen
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   2280
      Top             =   0
      Width           =   13335
   End
End
Attribute VB_Name = "frmVArrange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VR_No As Long
Dim R_Com_Code As String
Dim TDate As Date
Dim TMode As String

Private Sub cmbJoin_Click()
    If VR_No = 0 Then
        MsgBox "VR Number NOT Selected", vbExclamation
        Exit Sub
    End If
    PR_HRS_Open_CON
    HRS.Execute "INSERT INTO HRS_TR_Allocation(All_No,VR_No,V_No,ALL_UID) Values(0," & VR_No & ",'NA','" & UserID & "')"
    FN_Status_Update 4, VR_No, Module_ID, Sub_Module_ID
    Com_Code = Trim(MSFRequest.TextMatrix(MSFRequest.Row, 3))
    PR_FMT_Grid
    PR_FMT_Grid_Allocation
    PR_HRS_Close_CON
    VR_No = 0
End Sub

Private Sub cmdAllocate_Click()
On Error GoTo er_EH:
    Dim V_No, Std_Reason As String
    If Trim(DcmbVno.Text) = "" Or Trim(DcmbVno.Text) = "-Select-" Then
        MsgBox "Vehicle Number NOT Selected", vbExclamation
        Exit Sub
    Else
        V_No = Trim(DcmbVno.Text)
    End If
    
    If Trim(DcmbStdReason.Text) = "" Or Trim(DcmbStdReason.Text) = "-Select-" Then
        MsgBox "Standard Reason NOT Selected", vbExclamation
        Exit Sub
    Else
        V_No = Trim(DcmbVno.Text)
        Std_Reason = Trim(DcmbStdReason.Text)
        FN_Find_Reason_ID Std_Reason, Com_Code
    End If
    

    Dim All_No As Long
    Set RS = New ADODB.Recordset
    PR_HRS_Open_CON
    RS.Open "Select * from HRSV_TR_Allocations_WIP Where All_No=0 and All_UID='" & UserID & "'", HRS, adOpenStatic, adLockReadOnly
    If RS.EOF = False Then
        Set RS = New ADODB.Recordset
        RS.Open "SELECT (Max_All_No+1) as All_No FROM HRSV_TR_Max_Allocation_No", HRS, adOpenStatic, adLockReadOnly
        All_No = Val(RS!All_No)
        Set UPRS = New ADODB.Recordset
        UPRS.Open "Update HRS_TR_Allocation set All_No=" & All_No & ",V_No='" & Trim(DcmbVno.Text) & "', All_UID='" & UserID & "',All_Date_Time='" & Format(Date, "MM/dd/yyyy") + " " + Format(Time, "hh:mm:ss") & "',Reason_ID=" & Reason_ID & " Where All_No=0 and All_UID='" & UserID & "'", HRS, adOpenStatic, adLockOptimistic
        Dim LRS As ADODB.Recordset
        Set LRS = New ADODB.Recordset
        Dim SMSRS As ADODB.Recordset
        Dim Message As String
        LRS.Open "Select * from HRS_TR_Allocation Where All_No=" & All_No & " and All_UID='" & UserID & "'", HRS, adOpenStatic, adLockReadOnly
        Do While LRS.EOF = False
            FN_Status_Update 5, Val(LRS!VR_No), Module_ID, Sub_Module_ID
            Set SMSRS = New ADODB.Recordset
            PR_HRS_Open_CON
            SMSRS.Open "SELECT U_Contact1 FROM HRSV_TR_Requests Where VR_No=" & Val(LRS!VR_No), HRS, adOpenStatic, adLockReadOnly
            If SMSRS.EOF = False Then
                Dim VRS As ADODB.Recordset
                Set VRS = New ADODB.Recordset
                VRS.Open "Select * from HRSV_TR_MSTR_Vehicles_App Where V_No='" & Trim(DcmbVno.Text) & "'", HRS, adOpenStatic, adLockReadOnly
                If VRS.EOF = False Then
                    Message = Chr$(13) & "VR No.-" & Trim(LRS!VR_No) & " Arranged " & Chr$(13) & "Vehicle No." & Trim(DcmbVno.Text) & Chr$(13) & "Driver-" & Trim(VRS!D_Name) & Chr$(13) & "Contact No-" & Trim(VRS!D_Cont1) & " / " & Trim(VRS!D_cont2)
                End If
                FN_SMS Trim(SMSRS!U_Contact1), Message
            End If
            LRS.MoveNext
        Loop
        MsgBox "Allocation Number" & Str(All_No) & " Successfully Allocated", vbInformation
        FN_RPT_Allocation All_No
        PR_FMT_Grid_Allocation
    Else
        MsgBox "There is NO Records to allocate", vbExclamation
        Exit Sub
    End If
    PR_HRS_Close_CON
    Exit Sub
    
er_EH:
    MsgBox Err.Description, vbCritical
    Err.Clear
End Sub

Private Sub cmdChange_Click()
On Error GoTo er_EH:
    If MsgBox("Are you sure Do you Want to Change Allocated Vehicle ?", vbQuestion + vbYesNo, "Change Vehicle") = vbYes Then
        PR_HRS_Open_CON
        Set RS = New ADODB.Recordset
        RS.Open "Select * from HRS_sys_Request_Log Where Request_ID in (Select VR_No From HRS_V_Allocation Where All_No=" & Val(txtcAllNo.Text) & ") and status_ID=6", HRS, adOpenStatic, adLockReadOnly
        If RS.EOF = True Then
            Set UPRS = New ADODB.Recordset
            UPRS.Open "Update HRS_V_Allocation set V_No='" & Trim(dCmbCVNo.Text) & "' Where All_No=" & Val(txtcAllNo.Text), HRS, adOpenStatic, adLockOptimistic
            PR_HRS_Close_CON
            MsgBox "Allocated Vehicle Changed Successfully", vbInformation
        Else
            MsgBox "Vehicle Change Failed, VR Status is Closed", vbCritical
        End If
        Call cmdFind_Click
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

Private Sub cmdFind_Click()
    If Trim(txtcAllNo.Text) = "" Then
        MsgBox "Allocation Number NOT Entered", vbExclamation
        Exit Sub
    Else
        Dim AllNo As Long
        AllNo = Trim(txtcAllNo.Text)
        Set RS = New ADODB.Recordset
        PR_HRS_Open_CON
        RS.Open "Select * from HRSV_TR_Allocations Where All_No=" & AllNo, HRS, adOpenStatic, adLockReadOnly
        If RS.EOF = False Then
            lblCEDName.Caption = Trim(RS!D_Name)
            lblCEDCont.Caption = Trim(RS!Driver_Cont1) & " / " & Trim(RS!Driver_cont2)
            lblCEVCat.Caption = Trim(RS!Category)
            lblCEVModel.Caption = Trim(RS!Brand_Name) & " " & Trim(RS!Model_Name)
        Else
            lblCEDName.Caption = "-"
            lblCEDCont.Caption = "-"
            lblCEVCat.Caption = "-"
            lblCEVModel.Caption = "-"
        End If
    End If
End Sub

Private Sub cmdPrint_Click()
On Error GoTo er_EH:
    If Trim(txtAllNo.Text) = "" Then
        MsgBox "Allocation Numner NOT Entered", vbExclamation
        Exit Sub
    Else
        FN_RPT_Allocation Val(txtAllNo.Text)
    End If

    Exit Sub
er_EH:
    MsgBox Err.Description, vbCritical
    Err.Clear
End Sub

Private Sub cmdRefresh_Click()
    PR_FMT_Grid
End Sub

Private Sub cmdRemove_Click()
On Error GoTo er_EH:
    VR_No = Val(MSFAllocated.TextMatrix(MSFAllocated.Row, 0))
    If VR_No = 0 Then
        MsgBox "VR Number NOT Selected", vbExclamation
        Exit Sub
    End If
    PR_HRS_Open_CON
    HRS.Execute "DELETE FROM HRS_TR_Allocation WHERE ALL_NO=0 AND VR_NO=" & VR_No
    HRS.Execute "DELETE FROM HRS_sys_Request_Log Where status_ID=4 and REQUEST_ID=" & VR_No
    
    Set UPRS = New ADODB.Recordset
    UPRS.Open "Update HRS_TR_Request set Status_ID=2 Where VR_NO=" & VR_No, HRS, adOpenStatic, adLockOptimistic
    
    
    PR_FMT_Grid
    PR_FMT_Grid_Allocation
    PR_HRS_Close_CON
    VR_No = 0
    Exit Sub
er_EH:
    MsgBox Err.Description, vbCritical
    Err.Clear
End Sub
Private Sub dCmbCVNo_Click(Area As Integer)
    Set RS = New ADODB.Recordset
    PR_HRS_Open_CON
    RS.Open "Select * from HRSV_TR_MSTR_Vehicles_App Where V_No='" & Trim(dCmbCVNo.Text) & "' and Sub_Module_Name='Adhoc'", HRS, adOpenKeyset, adLockPessimistic
    If RS.EOF = False Then
        If IsNull(RS!D_Name) = False Then
            lblCNDName.Caption = Trim(RS!D_Name)
        Else
            lblCNDName.Caption = "-"
        End If
        
        If IsNull(RS!D_Cont1) = False Then
            lblCNDCont.Caption = Trim(RS!D_Cont1) & " / " & Trim(RS!D_cont2)
        Else
            lblCNDCont.Caption = "-"
        End If
        lblCNVCat.Caption = Trim(RS!Category)
        lblCNVModel.Caption = Trim(RS!Brand_Name) & " " & Trim(RS!Model_Name)
    Else
        lblCNDName.Caption = "-"
        lblCNDCont.Caption = "-"
        lblCNVCat.Caption = "-"
        lblCNVModel.Caption = "-"
    End If
    PR_HRS_Close_CON
End Sub

Private Sub DcmbVno_Click(Area As Integer)
On Error GoTo er_EH:
    PR_HRS_Open_CON
    Set RS = New ADODB.Recordset
    If Trim(DcmbVno.Text) <> "-Select-" Then
        RS.Open "Select * from HRSV_TR_MSTR_Vehicles_App Where V_No='" & Trim(DcmbVno.Text) & "'", HRS, adOpenStatic, adLockReadOnly
        If RS.EOF = False Then
            If IsNull(RS!D_Name) = True Then
                lblDriver.Caption = ""
            Else
                lblDriver.Caption = Trim(RS!D_Name)
            End If
            If IsNull(RS!D_Cont1) = True Then
                lblDno1.Caption = ""
            Else
                lblDno1.Caption = Trim(RS!D_Cont1)
            End If
            If IsNull(RS!D_cont2) = True Then
                lblDno2.Caption = ""
            Else
                lblDno2.Caption = Trim(RS!D_cont2)
            End If
            TMode = "Lic"
            TDate = RS!Lic_Exp
            PR_V_Andon
            TMode = "Ins"
            TDate = RS!Ins_Exp
            PR_V_Andon
            If Trim(RS!V_Grade) = "GREEN" Then
                lblGrade.BackColor = &H80FF80
            End If
            If Trim(RS!V_Grade) = "AMBER" Then
                lblGrade.BackColor = &H80C0FF
            End If
            If Trim(RS!V_Grade) = "RED" Then
                lblGrade.BackColor = &H8080FF
            End If
        Else
            lblDriver.Caption = ""
            lblDno1.Caption = ""
            lblDno2.Caption = ""
        End If
    End If

    PR_HRS_Close_CON
    Exit Sub
    
er_EH:
    MsgBox Err.Description, vbCritical
    Err.Clear
End Sub

Private Sub Form_Load()
    DTPTDate.Value = Format(Date, "MM/dd/yyyy")
    PR_FMT_Grid
    PR_FMT_Grid_Allocation
    PR_Fill_Vehicles
    Opt2.Value = True
    PR_Fill_Std_Reasons
End Sub

Public Sub PR_FMT_Grid()
On Error GoTo er_EH:
PR_HRS_Open_CON
        MSFRequest.Cols = 10
        Set GRDRS = New ADODB.Recordset
        R = 1
        MSFRequest.Clear
        GRDRS.Open "Select * from HRSV_TR_Requests Where convert(date,Req_Date_Time)='" & Format(DTPTDate.Value, "MM/dd/yyyy") & "' AND Status_ID=2 and Req_ComCode='" & U_Com_Code & "'", HRS, adOpenStatic, adLockReadOnly
        If GRDRS.EOF = False Then
            Do While GRDRS.EOF = False
                MSFRequest.Rows = R + 1
                MSFRequest.TextMatrix(R, 0) = Val(GRDRS!VR_No)
                MSFRequest.TextMatrix(R, 1) = Format(GRDRS!Req_Date_Time, "dd-MMM-yyyy")
                MSFRequest.TextMatrix(R, 2) = Format(GRDRS!Req_Date_Time, "HH:mm:ss")
                MSFRequest.TextMatrix(R, 3) = Trim(GRDRS!Req_ComCode)
                MSFRequest.TextMatrix(R, 4) = Trim(GRDRS!To_City)
                MSFRequest.TextMatrix(R, 5) = Trim(GRDRS!Loc_Dtls)
                MSFRequest.TextMatrix(R, 6) = Trim(GRDRS!Remarks)
                MSFRequest.TextMatrix(R, 7) = Trim(GRDRS!U_DName)
                MSFRequest.TextMatrix(R, 8) = Trim(GRDRS!Req_Passengers)
                MSFRequest.TextMatrix(R, 9) = Trim(GRDRS!Category)
                R = R + 1
                GRDRS.MoveNext
            Loop
        Else
            MSFRequest.Rows = 1
        End If
        
        MSFRequest.ColWidth(0) = 800
        MSFRequest.ColWidth(1) = 1000
        MSFRequest.ColWidth(2) = 1000
        MSFRequest.ColWidth(3) = 800
        MSFRequest.ColWidth(4) = 2000
        MSFRequest.ColWidth(5) = 2000
        MSFRequest.ColWidth(6) = 2000
        MSFRequest.ColWidth(7) = 2000
        MSFRequest.ColWidth(8) = 1200
        
        MSFRequest.TextMatrix(0, 0) = "VR NO."
        MSFRequest.TextMatrix(0, 1) = "REQ. DATE"
        MSFRequest.TextMatrix(0, 2) = "REQ. TIME"
        MSFRequest.TextMatrix(0, 3) = "FROM"
        MSFRequest.TextMatrix(0, 4) = "TO CITY"
        MSFRequest.TextMatrix(0, 5) = "LOCATION DETAILS"
        MSFRequest.TextMatrix(0, 6) = "REMARKS"
        MSFRequest.TextMatrix(0, 7) = "REQUESTER"
        MSFRequest.TextMatrix(0, 8) = "PASSENGERS"
        MSFRequest.TextMatrix(0, 9) = "CATEGORY"
    GRDRS.Close
    PR_HRS_Close_CON
    Exit Sub
    
er_EH:
    MsgBox Err.Description, vbCritical
    Err.Clear
End Sub
Private Sub MSFRequest_Click()
    VR_No = Val(MSFRequest.TextMatrix(MSFRequest.Row, 0))
    R_Com_Code = Trim(MSFRequest.TextMatrix(MSFRequest.Row, 3))
    MSFRequest.ToolTipText = Trim(MSFRequest.TextMatrix(MSFRequest.Row, MSFRequest.Col))
End Sub

Public Sub PR_FMT_Grid_Allocation()
On Error GoTo er_EH:
PR_HRS_Open_CON
        MSFAllocated.Cols = 10
        Set GRDRS = New ADODB.Recordset
        R = 1
        MSFAllocated.Clear
        GRDRS.Open "Select * from HRSV_TR_Requests Where Status_ID=4 and Req_ComCode='" & U_Com_Code & "'", HRS, adOpenStatic, adLockReadOnly
        If GRDRS.EOF = False Then
            Do While GRDRS.EOF = False
                MSFAllocated.Rows = R + 1
                MSFAllocated.TextMatrix(R, 0) = Val(GRDRS!VR_No)
                MSFAllocated.TextMatrix(R, 1) = Format(GRDRS!Req_Date_Time, "dd-MMM-yyyy")
                MSFAllocated.TextMatrix(R, 2) = Format(GRDRS!Req_Date_Time, "HH:mm:ss")
                MSFAllocated.TextMatrix(R, 3) = Trim(GRDRS!Req_ComCode)
                MSFAllocated.TextMatrix(R, 4) = Trim(GRDRS!To_City)
                MSFAllocated.TextMatrix(R, 5) = Trim(GRDRS!Loc_Dtls)
                MSFAllocated.TextMatrix(R, 6) = Trim(GRDRS!Remarks)
                MSFAllocated.TextMatrix(R, 7) = Trim(GRDRS!U_DName)
                MSFAllocated.TextMatrix(R, 8) = Trim(GRDRS!Req_Passengers)
                MSFAllocated.TextMatrix(R, 9) = Trim(GRDRS!Category)
                R = R + 1
                GRDRS.MoveNext
            Loop
        Else
            MSFAllocated.Rows = 1
        End If
        
        MSFAllocated.ColWidth(0) = 800
        MSFAllocated.ColWidth(1) = 1000
        MSFAllocated.ColWidth(2) = 1000
        MSFAllocated.ColWidth(3) = 800
        MSFAllocated.ColWidth(4) = 2000
        MSFAllocated.ColWidth(5) = 2000
        MSFAllocated.ColWidth(6) = 2000
        MSFAllocated.ColWidth(7) = 2000
        MSFAllocated.ColWidth(8) = 1200
        
        MSFAllocated.TextMatrix(0, 0) = "VR NO."
        MSFAllocated.TextMatrix(0, 1) = "REQ. DATE"
        MSFAllocated.TextMatrix(0, 2) = "REQ. TIME"
        MSFAllocated.TextMatrix(0, 3) = "FROM"
        MSFAllocated.TextMatrix(0, 4) = "TO CITY"
        MSFAllocated.TextMatrix(0, 5) = "REMARKS"
        MSFAllocated.TextMatrix(0, 6) = "LOCATION DETAILS"
        MSFAllocated.TextMatrix(0, 7) = "REQUESTER"
        MSFAllocated.TextMatrix(0, 8) = "PASSENGERS"
        MSFAllocated.TextMatrix(0, 9) = "CATEGORY"
        PR_HRS_Close_CON
    Exit Sub
er_EH:
    MsgBox Err.Description, vbCritical
    Err.Clear
End Sub

Public Sub PR_Fill_Vehicles()
    Set RS = New ADODB.Recordset
    PR_HRS_Open_CON
    RS.Open "Select * from HRSV_TR_MSTR_Vehicles_App Where Module_ID=1 and sub_Module_ID=1", HRS, adOpenStatic, adLockReadOnly
    DcmbVno.ListField = "V_No"
    Set DcmbVno.RowSource = RS
    dCmbCVNo.ListField = "V_No"
    Set dCmbCVNo.RowSource = RS
    DcmbVno.Text = "-Select-"
    dCmbCVNo.Text = "-Select-"
    PR_HRS_Close_CON
End Sub

Public Sub PR_V_Andon()
    
    If TDate < Date - 7 Then
        If TMode = "Lic" Then
            lblLic.BackColor = &H8080FF
        End If
        
        If TMode = "Ins" Then
            lblIns.BackColor = &H8080FF
        End If
    Else
        If TMode = "Lic" Then
            lblLic.BackColor = &H80FF80
        End If
        If TMode = "Ins" Then
            lblIns.BackColor = &H80FF80
        End If
    End If
End Sub

Public Sub PR_Fill_Std_Reasons()
    Set RS = New ADODB.Recordset
    PR_HRS_Open_CON
    RS.Open "Select Reason_Details from HRSV_TR_MSTR_Reason Order by Reason_Details", HRS, adOpenStatic, adLockReadOnly
    DcmbStdReason.ListField = "Reason_Details"
    Set DcmbStdReason.RowSource = RS
    DcmbStdReason.Text = "-Select-"
    PR_HRS_Close_CON
End Sub

Public Function FN_RPT_Allocation(ByRef AllNo As Long)
On Error GoTo er_EH:
    PR_REPORT_PATH
    CRW.ReportFileName = Report_Path + "HRS_V_RPT001.rpt"
    CRW.ParameterFields(0) = "AllocationNo;" & AllNo & "; true"
    CRW.Connect = DSN_SETTINGS
    If Opt1.Value = True Then
        CRW.Destination = crptToPrinter
    End If
    If Opt2.Value = True Then
        CRW.Destination = crptToWindow
    End If
    CRW.Action = 1
    Exit Function
    
er_EH:
    MsgBox Err.Description, vbCritical
    Err.Clear
End Function

