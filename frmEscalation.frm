VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmEscalation 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Requisition Escalation"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8145
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
   ScaleHeight     =   3435
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   2130
      Left            =   6960
      TabIndex        =   13
      Top             =   840
      Width           =   1095
      Begin VB.CommandButton cmdEscalation 
         Caption         =   "Escalate"
         Height          =   735
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
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
         Height          =   855
         Left            =   120
         Picture         =   "frmEscalation.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Exit from the System"
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton Command1 
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
         Picture         =   "frmEscalation.frx":0845
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Exit from the System"
         Top             =   7080
         Width           =   495
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   3060
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   9
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
      EndProperty
   End
   Begin VB.Label lblCost 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2400
      TabIndex        =   12
      Top             =   2640
      Width           =   4455
   End
   Begin VB.Label lblDistance 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2400
      TabIndex        =   11
      Top             =   2280
      Width           =   4455
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Approx. Cost (Rs.)"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Approx. Distance (Km)"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label lblRemarks 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   1920
      Width           =   4455
   End
   Begin VB.Label lblTo 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   1560
      Width           =   4455
   End
   Begin VB.Label lblFrom 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   1200
      Width           =   4455
   End
   Begin VB.Label lblVRNo 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Remarks"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "To"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "From"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "VR Number"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   2175
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   2280
      Picture         =   "frmEscalation.frx":108A
      Stretch         =   -1  'True
      Top             =   120
      Width           =   5175
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   0
      Picture         =   "frmEscalation.frx":33D6
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
      Left            =   2160
      Top             =   0
      Width           =   11175
   End
End
Attribute VB_Name = "frmEscalation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEscalation_Click()
    PR_HRS_Open_CON
    Set RS = New ADODB.Recordset
    RS.Open "UPDATE HRS_TR_Request SET STATUS_ID=7 WHERE VR_NO=" & Val(lblVRNo.Caption), HRS, adOpenStatic, adLockOptimistic
    PR_HRS_Close_CON
    frmVReq.PR_Grid_My_Request
    MsgBox "VR Escalated Successfully", vbInformation
    Close All
    Unload Me
End Sub

Private Sub cmdExit_Click()
    Close All
    Unload Me
End Sub

Private Sub DcmbRPerson_Change()
    PR_R_Person
End Sub

Private Sub Form_Load()
    PR_Initialization
End Sub

Public Sub PR_Initialization()
    Set RS = New ADODB.Recordset
    PR_HRS_Open_CON
    RS.Open "Select * from HRSV_TR_Requests Where VR_No='" & Tmp & "'", HRS, adOpenStatic, adLockReadOnly
    If RS.EOF = False Then
        lblVRNo.Caption = Tmp
        lblFrom.Caption = Trim(RS!Req_ComCode)
        lblTo.Caption = Trim(RS!To_City)
        lblRemarks.Caption = Trim(RS!Remarks)
        lblDistance.Caption = Trim(RS!Distance)
        lblCost.Caption = Format(Val(RS!Distance) * Val(RS!AC_Rate), "###,#00.00")
    End If
    
'    Set RS = New ADODB.Recordset
'    RS.Open "Select * from HRSV_sys_Special_Authority Order By U_DName", HRS, adOpenStatic, adLockReadOnly
'    DcmbRPerson.ListField = "U_DName"
'    Set DcmbRPerson.RowSource = RS
'    DcmbRPerson.Text = "-Select-"
    PR_HRS_Close_CON
End Sub

Public Sub PR_R_Person()
    Set RS = New ADODB.Recordset
    PR_HRS_Open_CON
    RS.Open "Select * from HRSV_sys_SP_Approval Where U_DName='" & Trim(DcmbRPerson.Text) & "'", HRS, adOpenStatic, adLockReadOnly
    If RS.EOF = False Then
        lblRmail.Caption = Trim(RS!U_email)
        lblRCont1.Caption = Trim(RS!U_Contact1)
    Else
        lblRmail.Caption = ""
        lblRCont1.Caption = ""
    End If
    PR_HRS_Close_CON
End Sub
