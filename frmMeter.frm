VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMeter 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Meter Readings"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13020
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
   ScaleHeight     =   8310
   ScaleWidth      =   13020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   7215
      Left            =   12240
      TabIndex        =   7
      Top             =   750
      Width           =   735
      Begin VB.CommandButton cmdExit 
         Height          =   495
         Left            =   120
         Picture         =   "frmMeter.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Exit from the System"
         Top             =   6600
         Width           =   495
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   6
      Top             =   7950
      Width           =   13020
      _ExtentX        =   22966
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   14
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "10/17/2020"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "12:53 AM"
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
         BeginProperty Panel12 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel13 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel14 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6975
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   12303
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "&Requisition Breakdown"
      TabPicture(0)   =   "frmMeter.frx":0845
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblCKm"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblCost"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblKm"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label8"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label7"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label6"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label5"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblDriverContact"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblDriverName"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblVno"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label2"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "MSFRequest"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtIn"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtOut"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cmdStop"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Picture1"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtAllNo"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "cmdFind"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Picture3"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).ControlCount=   21
      TabCaption(1)   =   "&Allocation Finalization"
      TabPicture(1)   =   "frmMeter.frx":0861
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture2"
      Tab(1).Control(1)=   "cmdApprove"
      Tab(1).Control(2)=   "cmdDelete"
      Tab(1).Control(3)=   "txtAPDeduction"
      Tab(1).Control(4)=   "cmdAdd"
      Tab(1).Control(5)=   "txtAPPayment"
      Tab(1).Control(6)=   "txtAPReason"
      Tab(1).Control(7)=   "MSFAPAllocation"
      Tab(1).Control(8)=   "Opt2"
      Tab(1).Control(9)=   "Opt1"
      Tab(1).Control(10)=   "cmdRefresh"
      Tab(1).Control(11)=   "DTPTDate"
      Tab(1).Control(12)=   "MSFAllocations"
      Tab(1).Control(13)=   "lblAPVNo"
      Tab(1).Control(14)=   "lblAllNo"
      Tab(1).Control(15)=   "Label17"
      Tab(1).Control(16)=   "Label16"
      Tab(1).Control(17)=   "Label14"
      Tab(1).Control(18)=   "Label13"
      Tab(1).Control(19)=   "Label12"
      Tab(1).Control(20)=   "Label11"
      Tab(1).Control(21)=   "Label10"
      Tab(1).Control(22)=   "Label9"
      Tab(1).Control(23)=   "Label3"
      Tab(1).ControlCount=   24
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   6240
         Picture         =   "frmMeter.frx":087D
         ScaleHeight     =   1215
         ScaleWidth      =   5655
         TabIndex        =   51
         Top             =   120
         Width           =   5655
      End
      Begin VB.PictureBox Picture2 
         Height          =   495
         Left            =   -69600
         ScaleHeight     =   435
         ScaleWidth      =   6435
         TabIndex        =   49
         Top             =   6100
         Width           =   6495
         Begin VB.CommandButton cmdConfirmAllocation 
            Caption         =   "&Process"
            Height          =   375
            Left            =   4920
            TabIndex        =   50
            Top             =   30
            Width           =   1455
         End
      End
      Begin VB.CommandButton cmdApprove 
         Caption         =   "Approve All"
         Height          =   375
         Left            =   -65400
         TabIndex        =   46
         Top             =   120
         Width           =   2295
      End
      Begin VB.CommandButton cmdDelete 
         Height          =   615
         Left            =   -70200
         Picture         =   "frmMeter.frx":4931
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   5400
         Width           =   495
      End
      Begin VB.TextBox txtAPDeduction 
         Height          =   285
         Left            =   -72240
         TabIndex        =   42
         Top             =   5400
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Height          =   615
         Left            =   -71040
         Picture         =   "frmMeter.frx":4E37
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   5400
         Width           =   735
      End
      Begin VB.TextBox txtAPPayment 
         Height          =   285
         Left            =   -73560
         TabIndex        =   36
         Top             =   5400
         Width           =   1215
      End
      Begin VB.TextBox txtAPReason 
         Height          =   285
         Left            =   -73560
         TabIndex        =   34
         Top             =   5040
         Width           =   3855
      End
      Begin MSFlexGridLib.MSFlexGrid MSFAPAllocation 
         Height          =   1815
         Left            =   -69600
         TabIndex        =   31
         Top             =   4200
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   3201
         _Version        =   393216
      End
      Begin VB.OptionButton Opt2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Pending Vehicle Requisitions"
         Height          =   255
         Left            =   -74880
         TabIndex        =   30
         Top             =   120
         Width           =   2415
      End
      Begin VB.OptionButton Opt1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Transaction Date"
         Height          =   255
         Left            =   -72360
         TabIndex        =   29
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton cmdRefresh 
         Height          =   400
         Left            =   -68400
         Picture         =   "frmMeter.frx":543E
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   120
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPTDate 
         Height          =   375
         Left            =   -70560
         TabIndex        =   27
         Top             =   120
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         Format          =   55443457
         CurrentDate     =   43329
      End
      Begin VB.CommandButton cmdFind 
         Height          =   255
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "View Requisition Details"
         Top             =   120
         Width           =   855
      End
      Begin VB.TextBox txtAllNo 
         Height          =   285
         Left            =   2280
         TabIndex        =   0
         Top             =   120
         Width           =   1575
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00400000&
         Height          =   615
         Left            =   9720
         ScaleHeight     =   555
         ScaleWidth      =   2115
         TabIndex        =   24
         Top             =   6000
         Width           =   2175
         Begin VB.CommandButton cmdConfirm 
            Caption         =   "&Confirm Allocation"
            Height          =   375
            Left            =   120
            TabIndex        =   4
            Top             =   110
            Width           =   1935
         End
      End
      Begin VB.CommandButton cmdStop 
         Height          =   975
         Left            =   3960
         Picture         =   "frmMeter.frx":5CE4
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtOut 
         Height          =   285
         Left            =   2280
         TabIndex        =   2
         Top             =   1560
         Width           =   2535
      End
      Begin VB.TextBox txtIn 
         Height          =   285
         Left            =   2280
         TabIndex        =   3
         Top             =   1920
         Width           =   2535
      End
      Begin MSFlexGridLib.MSFlexGrid MSFRequest 
         Height          =   3615
         Left            =   120
         TabIndex        =   11
         Top             =   2400
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   6376
         _Version        =   393216
         Cols            =   5
      End
      Begin MSFlexGridLib.MSFlexGrid MSFAllocations 
         Height          =   3495
         Left            =   -74880
         TabIndex        =   25
         Top             =   600
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   6165
         _Version        =   393216
         Cols            =   5
      End
      Begin VB.Label lblAPVNo 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   -72960
         TabIndex        =   48
         Top             =   4560
         Width           =   1815
      End
      Begin VB.Label lblAllNo 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   -74880
         TabIndex        =   47
         Top             =   4560
         Width           =   1815
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Deductions"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   -72240
         TabIndex        =   44
         Top             =   5760
         Width           =   1095
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Payments"
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   -73560
         TabIndex        =   43
         Top             =   5760
         Width           =   1215
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   -72240
         TabIndex        =   40
         Top             =   6120
         Width           =   1095
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   -71040
         TabIndex        =   39
         Top             =   6120
         Width           =   1335
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   -73560
         TabIndex        =   38
         Top             =   6120
         Width           =   1215
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   -74880
         TabIndex        =   37
         Top             =   6120
         Width           =   1215
      End
      Begin VB.Label Label10 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Amount (Rs.)"
         Height          =   255
         Left            =   -74880
         TabIndex        =   35
         Top             =   5400
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Reason"
         Height          =   255
         Left            =   -74880
         TabIndex        =   33
         Top             =   5040
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Additional Payments"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   32
         Top             =   4200
         Width           =   5175
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Allocation Number"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   120
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Vehicle Number"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label lblVno 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   2280
         TabIndex        =   22
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label lblDriverName 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   840
         Width           =   3735
      End
      Begin VB.Label lblDriverContact 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1200
         Width           =   3735
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Out Meter"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "In Meter"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Total Distance (Km)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6240
         TabIndex        =   17
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Travel Cost (Rs.)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10080
         TabIndex        =   16
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label lblKm 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   495
         Left            =   6240
         TabIndex        =   15
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label lblCost 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   495
         Left            =   10080
         TabIndex        =   14
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label lblCKm 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   495
         Left            =   8160
         TabIndex        =   13
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cum.. Distance (Km)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8160
         TabIndex        =   12
         Top             =   1440
         Width           =   1815
      End
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Allocation Finalizations"
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
      TabIndex        =   5
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
      Width           =   13575
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   0
      Picture         =   "frmMeter.frx":6639
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "frmMeter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AC_Rate, NAC_Rate As Double
Dim Km As Double

Private Sub cmdAdd_Click()
On Error GoTo er_EH:
    If Trim(lblAllNo.Caption) = "" Then
        MsgBox "Allocation Number NOT Selected", vbExclamation
        Exit Sub
    End If
    
    If Trim(txtAPReason.Text) = "" Then
        MsgBox "Reason NOT Entered", vbExclamation
        Exit Sub
    End If
    
    If Trim(txtAPPayment.Text) = "" Then
        MsgBox "Payment Amount NOT Entered", vbExclamation
        Exit Sub
    End If
    
    Dim AllNo As Long
    Dim APReason As String
    Dim APPayment, APDeduction As Double
    
    AllNo = Val(lblAllNo.Caption)
    APReason = Trim(txtAPReason.Text)
    APPayment = Val(txtAPPayment.Text)
    APDeduction = Val(txtAPDeduction.Text)
    
    PR_HRS_Open_CON
    HRS.Execute "INSERT INTO HRS_TR_Other_Costs(All_No,Reason,Payment,Deduction,T_Date_Time,T_UID) " _
                + "VALUES(" & AllNo & ",'" & APReason & "'," & APPayment & "," & APDeduction & ",'" & Date + Time & "','" & UserID & "')"
    PR_FMT_AP AllNo
    PR_HRS_Close_CON
    Exit Sub
    
er_EH:
    MsgBox Err.Description, vbCritical
    Err.Clear
End Sub

Private Sub cmdApprove_Click()
On Error GoTo er_EH:
    If MsgBox("Are you sure Do you wish to APROVE All Selected Allocations ?", vbQuestion + vbYesNo, "Approve Record") = vbYes Then
        PR_Approve_all
        MsgBox "Allocations Confirmed Successfully", vbInformation
    End If
    PR_FMT_Allocations
    Exit Sub
er_EH:
    MsgBox Err.Description, vbCritical
    Err.Clear
End Sub

Private Sub cmdConfirm_Click()
    Dim AllNo As Long
    If Trim(txtAllNo.Text) = "" Then
        MsgBox "Allocation Number NOT Entered", vbExclamation
        txtAllNo.SetFocus
        Exit Sub
    End If
    
    If Trim(txtIn.Text) = "" Or Trim(txtOut.Text) = "" Then
        MsgBox "Meter Readings NOT Entered", vbExclamation
        Exit Sub
    End If
    PR_Cal
    
    If Val(lblCKm.Caption) = 0 Then
        MsgBox "Requisition Distance NOT Entered", vbExclamation
        Exit Sub
    End If
    
    If Val(lblKm.Caption) = Val(lblCKm.Caption) Then
        AllNo = Val(txtAllNo.Text)
        Dim T_AC_Km, T_NAC_Km As Double
        Dim T_AC_Cost, T_NAC_Cost As Double
        PR_HRS_Open_CON
        For I = 1 To MSFRequest.Rows - 1
            T_AC_Km = Val(MSFRequest.TextMatrix(I, 4))
            T_NAC_Km = Val(MSFRequest.TextMatrix(I, 5))
            T_AC_Cost = Val(MSFRequest.TextMatrix(I, 4)) * AC_Rate
            T_NAC_Cost = Val(MSFRequest.TextMatrix(I, 5)) * NAC_Rate
            PR_HRS_Open_CON
            Set RS = New ADODB.Recordset
            RS.Open "Select * from HRSV_TR_Requests Where VR_No=" & Val(MSFRequest.TextMatrix(I, 0)), HRS, adOpenStatic, adLockReadOnly
            If RS.EOF = False Then
                FN_Status_Update 6, Val(MSFRequest.TextMatrix(I, 0)), Module_ID, Sub_Module_ID
                PR_HRS_Open_CON
                Set UPRS = New ADODB.Recordset
                UPRS.Open "update HRS_TR_Request set V_No='" & Trim(lblVno.Caption) & "',AC_Km=" & T_AC_Km & ",NAC_Km=" & T_NAC_Km & ",AC_Cost=" & T_AC_Cost & ",NAC_Cost=" & T_NAC_Cost & " Where VR_NO=" & Val(MSFRequest.TextMatrix(I, 0)), HRS, adOpenStatic, adLockReadOnly
            Else
                MsgBox "VR Confirmation Failed", vbCritical
                Exit Sub
            End If
        Next I
        PR_Update_All_Header AllNo
        Call cmdFind_Click
        MsgBox "Allocation Confirmed Successfully", vbInformation
        PR_Initialization
        PR_HRS_Close_CON
    Else
        MsgBox "Meter Reading Mismatch with Cummulative Travelled Cost", vbExclamation
    End If
    PR_HRS_Close_CON
End Sub

Private Sub cmdConfirmAllocation_Click()
    If Trim(lblAllNo.Caption) = "" Then
        MsgBox "Allocation Number NOT Selected", vbExclamation
        Exit Sub
    End If
    Dim All_No As Long
    All_No = Val(lblAllNo.Caption)
    PR_HRS_Open_CON
    Set UPRS = New ADODB.Recordset
    UPRS.Open "Update HRS_TR_Other_Costs set Confirmed=1 Where All_No=" & All_No, HRS, adOpenStatic, adLockOptimistic
    PR_Update_Other_Cost All_No
    MsgBox "Other Payments Processed Successfully", vbInformation
    PR_FMT_Allocations
    PR_FMT_AP 0
    PR_HRS_Close_CON
End Sub

Private Sub cmdDelete_Click()
On Error GoTo er_EH:
    If MsgBox("Are you sure Do you Want to Delete your selected Payment ?", vbQuestion + vbYesNo, "Delete Record") = vbYes Then
        PR_HRS_Open_CON
        Dim All_No As Long
        All_No = Val(MSFAPAllocation.TextMatrix(MSFAPAllocation.Row, 0))
        HRS.Execute "Delete from HRS_TR_Other_Costs Where All_No=" & All_No & " and Reason='" & Trim(MSFAPAllocation.TextMatrix(MSFAPAllocation.Row, 1)) & "'"
        PR_HRS_Close_CON
        PR_FMT_AP All_No
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
    If Trim(txtAllNo.Text) = "" Then
        MsgBox "Allocation Numner NOT Enter", vbExclamation
        Exit Sub
    Else
        Screen.MousePointer = vbHourglass
        FN_FMT_GRID Val(txtAllNo.Text)
        Set RS = New ADODB.Recordset
        PR_HRS_Open_CON
        RS.Open "Select * from HRSV_TR_MSTR_Vehicles_App Where V_No='" & Trim(lblVno.Caption) & "'", HRS, adOpenStatic, adLockReadOnly
        If RS.EOF = False Then
            AC_Rate = Val(RS!AC_Rate)
            NAC_Rate = Val(RS!Non_AC_Rate)
        Else
            AC_Rate = 0
            NAC_Rate = 0
        End If
        PR_HRS_Close_CON
        Screen.MousePointer = vbDefault
    End If
End Sub

Public Function FN_FMT_GRID(ByRef AllNo As Long)
    Set GRDRS = New ADODB.Recordset
    PR_HRS_Open_CON
    GRDRS.Open "Select * from HRSV_TR_Allocations Where All_No=" & AllNo, HRS, adOpenStatic, adLockReadOnly
    MSFRequest.Cols = 7
    MSFRequest.Rows = 1
    R = 1
    If GRDRS.EOF = False Then
        lblDriverName.Caption = Trim(GRDRS!D_Name)
        lblDriverContact.Caption = Trim(GRDRS!Driver_Cont1) + " / " + Trim(GRDRS!Driver_Cont1)
        lblVno.Caption = Trim(GRDRS!V_No)
        
        Do While GRDRS.EOF = False
            MSFRequest.Rows = R + 1
            MSFRequest.TextMatrix(R, 0) = Trim(GRDRS!VR_No)
            MSFRequest.TextMatrix(R, 1) = Trim(GRDRS!U_DName)
            MSFRequest.TextMatrix(R, 2) = Trim(GRDRS!To_City)
            If IsNull(GRDRS!AC_Km) = True Then
                MSFRequest.TextMatrix(R, 4) = ""
            Else
                MSFRequest.TextMatrix(R, 4) = Trim(GRDRS!AC_Km)
            End If
            
            If IsNull(GRDRS!NAC_Km) = True Then
                MSFRequest.TextMatrix(R, 5) = ""
            Else
                MSFRequest.TextMatrix(R, 5) = Trim(GRDRS!NAC_Km)
            End If
                
            R = R + 1
            GRDRS.MoveNext
        Loop
    Else
        lblDriverName.Caption = "-"
        lblDriverContact.Caption = "-"
        lblVno.Caption = "-"
    End If
    MSFRequest.TextMatrix(0, 0) = "REQ.#"
    MSFRequest.TextMatrix(0, 1) = "REQUESTER"
    MSFRequest.TextMatrix(0, 2) = "MAIN CITY"
    MSFRequest.TextMatrix(0, 3) = "LOC. REMARKS"
    MSFRequest.TextMatrix(0, 4) = "AC(Km)"
    MSFRequest.TextMatrix(0, 5) = "NONE AC(Km)"
    MSFRequest.TextMatrix(0, 6) = "COST(Rs.)"
    
    MSFRequest.ColWidth(0) = 800
    MSFRequest.ColWidth(1) = 2000
    MSFRequest.ColWidth(2) = 2000
    MSFRequest.ColWidth(3) = 2000
    MSFRequest.ColWidth(4) = 1100
    MSFRequest.ColWidth(5) = 1100
    MSFRequest.ColWidth(6) = 1100
    PR_HRS_Close_CON
End Function

Private Sub cmdRefresh_Click()
    PR_FMT_Allocations
End Sub

Private Sub Form_Load()
    Module_ID = 1
    Sub_Module_ID = 1
    FN_FMT_GRID 0
    Opt2.Value = True
    DTPTDate.Value = Date
    PR_FMT_Allocations
    PR_FMT_AP 0
End Sub

Public Sub PR_Cal()
On Error GoTo er_EH:

    Exit Sub
    
er_EH:
    MsgBox Err.Description, vbCritical
    Err.Clear
End Sub

Private Sub MSFAllocations_Click()
    If MSFAllocations.Rows > 1 Then
        Set RS = New ADODB.Recordset
        PR_HRS_Open_CON
        RS.Open "SELECT * FROM HRSV_TR_Other_Costs WHERE ALL_NO<>" & Val(MSFAllocations.TextMatrix(MSFAllocations.Row, 0)) & " AND T_UID='" & UserID & "'", HRS, adOpenStatic, adLockReadOnly
        If RS.EOF = False Then
            MsgBox "Un Processed Other Costs records found, Please Process Or Delete found records", vbExclamation
            Exit Sub
        Else
            lblAllNo.Caption = Trim(MSFAllocations.TextMatrix(MSFAllocations.Row, 0))
            lblAPVNo.Caption = Trim(MSFAllocations.TextMatrix(MSFAllocations.Row, 1))
        End If
        PR_HRS_Close_CON
    End If
End Sub

Private Sub MSFRequest_KeyPress(KeyAscii As Integer)
On Error GoTo er_EH:
    If Trim(lblKm.Caption) <> "" Then
        If MSFRequest.Col = 4 Or MSFRequest.Col = 5 Then
            If KeyAscii = 22 Then
                MSFRequest.Clip = Clipboard.GetText
            End If
        
            If IsNumeric(Chr$(KeyAscii)) Then
                MSFRequest.Clip = MSFRequest.TextMatrix(MSFRequest.Row, MSFRequest.Col) + Chr$(KeyAscii)
            End If
        
            If KeyAscii = 8 Then
                MSFRequest.TextMatrix(MSFRequest.Row, MSFRequest.Col) = ""
            End If
            FN_Travel_Cost Trim(lblVno.Caption)
        End If
        If MSFRequest.Col = 7 Then
            If KeyAscii = 22 Then
                MSFRequest.Clip = Clipboard.GetText
            End If
            MSFRequest.Clip = MSFRequest.TextMatrix(MSFRequest.Row, MSFRequest.Col) + Chr$(KeyAscii)
            If KeyAscii = 8 Then
                MSFRequest.TextMatrix(MSFRequest.Row, MSFRequest.Col) = ""
            End If
        End If
    End If
Exit Sub
er_EH:
    MsgBox Err.Description, vbCritical
    Err.Clear
End Sub

Private Sub Opt1_Click()
    If Opt2.Value = True Then
        DTPTDate.Enabled = False
    Else
        DTPTDate.Enabled = True
    End If
End Sub

Private Sub Opt2_Click()
    If Opt1.Value = True Then
        DTPTDate.Enabled = True
    Else
        DTPTDate.Enabled = False
    End If
End Sub

Private Sub txtIn_Change()
    If Val(txtIn.Text) > 0 Then
        PR_Cal_Meter
    End If
End Sub

Private Sub txtOut_Change()
    If Val(txtOut.Text) > 0 Then
        PR_Cal_Meter
    End If
End Sub
Public Sub PR_Cal_Meter()
    Km = Val(txtIn.Text) - Val(txtOut.Text)
    If Val(txtIn.Text) > 0 Then
        If Val(txtOut.Text) > 0 Then
            If Km > 0 Then
                lblKm.Caption = Km
            Else
                lblKm.Caption = 0
            End If
        End If
    End If
End Sub

Public Function FN_Travel_Cost(ByRef VNo As String)
    Dim TCost As Double
    Dim SKM As Double
    TCost = 0
    For I = 1 To MSFRequest.Rows - 1
        SKM = SKM + (Val(MSFRequest.TextMatrix(I, 4)) + Val(MSFRequest.TextMatrix(I, 5)))
        If Km < SKM Then
            MsgBox "Total KM Exceeded", vbExclamation
            MSFRequest.TextMatrix(MSFRequest.Row, MSFRequest.Col) = ""
            Exit Function
        End If
        TCost = TCost + Val(MSFRequest.TextMatrix(I, 4)) * AC_Rate + Val(MSFRequest.TextMatrix(I, 5)) * NAC_Rate
        MSFRequest.TextMatrix(I, 6) = Val(MSFRequest.TextMatrix(I, 4)) * AC_Rate + Val(MSFRequest.TextMatrix(I, 5)) * NAC_Rate
        lblCost.Caption = Format(TCost, "##.00")
    Next I
    lblCKm.Caption = SKM
End Function

Public Sub PR_Update_All_Header(ByRef AllNo As Long)
    Dim cmd As ADODB.Command
    Dim RS1 As New ADODB.Recordset
    Set RS1 = New ADODB.Recordset
    Set cmd = New ADODB.Command
    Dim strconnect As String

    PR_HRS_Open_CON
    cmd.ActiveConnection = HRS
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "HRS_SP_MSTR_Allocation"
    cmd.Parameters.Append cmd.CreateParameter("@All_No", adInteger, adParamInput, 1, AllNo)
    cmd.Parameters.Append cmd.CreateParameter("@In_Meter", adInteger, adParamInput, 1, Val(txtIn.Text))
    cmd.Parameters.Append cmd.CreateParameter("@Out_Meter", adInteger, adParamInput, 1, Val(txtOut.Text))
    cmd.CommandTimeout = 0
    Set RS1 = cmd.Execute
    Set cmd.ActiveConnection = Nothing
End Sub

Public Sub PR_Initialization()
    txtIn.Text = ""
    txtOut.Text = ""
    lblKm.Caption = ""
    lblCKm.Caption = ""
    lblCost.Caption = ""
    txtAllNo.Text = 0
End Sub

Public Sub PR_FMT_Allocations()
On Error GoTo er_EH:
    Dim Other_Cost As Double
    PR_HRS_Open_CON
    Set GRDRS = New ADODB.Recordset
    If Opt2.Value = True Then
        GRDRS.Open "SELECT * FROM HRS_TR_MSTR_Allocation Where Confirmed=0", HRS, adOpenStatic, adLockReadOnly
    End If
    If Opt1.Value = True Then
        GRDRS.Open "SELECT * FROM HRS_TR_MSTR_Allocation Where Confirmed=0 AND CONVERT(DATE,T_DATE_TIME)='" & DTPTDate.Value & "'", HRS, adOpenStatic, adLockReadOnly
    End If
    
    MSFAllocations.Cols = 11
    MSFAllocations.Rows = 1
    R = 1
    Do While GRDRS.EOF = False
        MSFAllocations.Rows = R + 1
        MSFAllocations.TextMatrix(R, 0) = Trim(GRDRS!All_No)
        MSFAllocations.TextMatrix(R, 1) = Trim(GRDRS!V_No)
        MSFAllocations.TextMatrix(R, 2) = Trim(GRDRS!Out_Meter)
        MSFAllocations.TextMatrix(R, 3) = Trim(GRDRS!In_Meter)
        MSFAllocations.TextMatrix(R, 4) = Trim(GRDRS!AC_Km)
        MSFAllocations.TextMatrix(R, 5) = Trim(GRDRS!NAC_Km)
        MSFAllocations.TextMatrix(R, 6) = Val(GRDRS!AC_Km) + Val(GRDRS!NAC_Km)
        MSFAllocations.TextMatrix(R, 7) = Format(Val(GRDRS!NAC_Km) * Val(GRDRS!NAC_Rate), "##.00")
        MSFAllocations.TextMatrix(R, 8) = Format(Val(GRDRS!AC_Km) * Val(GRDRS!AC_Rate), "##.00")
        If IsNull(GRDRS!Other_Cost) = False Then
            MSFAllocations.TextMatrix(R, 9) = Format(Val(GRDRS!Other_Cost), "##.00")
            Other_Cost = Format(Val(GRDRS!Other_Cost), "##.00")
        Else
            MSFAllocations.TextMatrix(R, 9) = 0
            Other_Cost = 0
        End If
        MSFAllocations.TextMatrix(R, 10) = Format(Val(GRDRS!NAC_Km) * Val(GRDRS!NAC_Rate) + Val(GRDRS!AC_Km) * Val(GRDRS!AC_Rate) + Other_Cost, "##.00")
        R = R + 1
        GRDRS.MoveNext
    Loop
    
    MSFAllocations.TextMatrix(0, 0) = "ALL#"
    MSFAllocations.TextMatrix(0, 1) = "VEHI#"
    MSFAllocations.TextMatrix(0, 2) = "OUT METER"
    MSFAllocations.TextMatrix(0, 3) = "IN METER"
    MSFAllocations.TextMatrix(0, 4) = "AC KM"
    MSFAllocations.TextMatrix(0, 5) = "NAC KM"
    MSFAllocations.TextMatrix(0, 6) = "TOTAL KM"
    MSFAllocations.TextMatrix(0, 7) = "NAC COST"
    MSFAllocations.TextMatrix(0, 8) = "AC COST"
    MSFAllocations.TextMatrix(0, 9) = "OTHER COST"
    MSFAllocations.TextMatrix(0, 10) = "TOTAL COST"
    
    MSFAllocations.ColWidth(0) = 800
    MSFAllocations.ColWidth(1) = 1000
    MSFAllocations.ColWidth(2) = 1000
    MSFAllocations.ColWidth(3) = 1000
    MSFAllocations.ColWidth(4) = 1000
    MSFAllocations.ColWidth(5) = 1000
    MSFAllocations.ColWidth(6) = 1000
    MSFAllocations.ColWidth(7) = 1000
    MSFAllocations.ColWidth(8) = 1000
    MSFAllocations.ColWidth(9) = 1000
    MSFAllocations.ColWidth(10) = 1000
    
    PR_HRS_Close_CON
    Exit Sub
    
er_EH:
    MsgBox Err.Description, vbCritical
    Err.Clear
End Sub

Public Sub PR_FMT_AP(ByRef PRAllNo As Long)
    PR_HRS_Open_CON
    Set GRDRS = New ADODB.Recordset
    If PRAllNo = 0 Then
        GRDRS.Open "Select * from HRSV_TR_Other_Costs Where Confirmed=0", HRS, adOpenStatic, adLockReadOnly
    Else
        GRDRS.Open "Select * from HRSV_TR_Other_Costs Where All_no=" & PRAllNo & "and Confirmed=0", HRS, adOpenStatic, adLockReadOnly
    End If
    R = 1
    MSFAPAllocation.Rows = 1
    MSFAPAllocation.Cols = 4
    Do While GRDRS.EOF = False
        MSFAPAllocation.Rows = R + 1
        MSFAPAllocation.TextMatrix(R, 0) = Trim(GRDRS!All_No)
        MSFAPAllocation.TextMatrix(R, 1) = Trim(GRDRS!Reason)
        MSFAPAllocation.TextMatrix(R, 2) = Trim(GRDRS!Payment)
        MSFAPAllocation.TextMatrix(R, 3) = Trim(GRDRS!Deduction)
        R = R + 1
        GRDRS.MoveNext
    Loop
    
    MSFAPAllocation.TextMatrix(0, 0) = "ALL#"
    MSFAPAllocation.TextMatrix(0, 1) = "REASON"
    MSFAPAllocation.TextMatrix(0, 2) = "PAYMENT"
    MSFAPAllocation.TextMatrix(0, 3) = "DEDUCTION"
    
    MSFAPAllocation.ColWidth(0) = 800
    MSFAPAllocation.ColWidth(1) = 2000
    MSFAPAllocation.ColWidth(2) = 1000
    MSFAPAllocation.ColWidth(3) = 1000

    PR_HRS_Close_CON
End Sub

Public Sub PR_Update_Other_Cost(ByRef AllNo As Long)
    Dim cmd As ADODB.Command
    Dim RS1 As New ADODB.Recordset
    Set RS1 = New ADODB.Recordset
    Set cmd = New ADODB.Command
    Dim strconnect As String

    PR_HRS_Open_CON
    cmd.ActiveConnection = HRS
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "HRS_SP_Other_Cost"
    cmd.Parameters.Append cmd.CreateParameter("@All_No", adInteger, adParamInput, 1, AllNo)
    cmd.CommandTimeout = 0
    Set RS1 = cmd.Execute
    Set cmd.ActiveConnection = Nothing
End Sub

Public Sub PR_Approve_all()
    Dim cmd As ADODB.Command
    Dim RS1 As New ADODB.Recordset
    Set RS1 = New ADODB.Recordset
    Set cmd = New ADODB.Command
    Dim strconnect As String

    PR_HRS_Open_CON
    cmd.ActiveConnection = HRS
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "HRS_SP_Confirm_Allocations"
    Dim Date_Flag As Integer
    If Opt1.Value = True Then
        Date_Flag = 0
    Else
        Date_Flag = 1
    End If
    Dim T_Date As Date
    T_Date = DTPTDate.Value
    cmd.Parameters.Append cmd.CreateParameter("@User_ID", adChar, adParamInput, 20, UserID)
    cmd.Parameters.Append cmd.CreateParameter("@Date_Flag", adInteger, adParamInput, 1, Date_Flag)
    cmd.Parameters.Append cmd.CreateParameter("@Trans_Date", adDate, adParamInput, 8, T_Date)
    cmd.CommandTimeout = 0
    Set RS1 = cmd.Execute
    Set cmd.ActiveConnection = Nothing
End Sub
