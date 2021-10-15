VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmVReq 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Vehicle Requisition"
   ClientHeight    =   9450
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12210
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
   ScaleHeight     =   9450
   ScaleWidth      =   12210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   8175
      Left            =   120
      TabIndex        =   14
      Top             =   840
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   14420
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "&Vehicle Requisition"
      TabPicture(0)   =   "frmVReq.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&My Requisitions"
      TabPicture(1)   =   "frmVReq.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdMyReq"
      Tab(1).Control(1)=   "MSFMYRequest"
      Tab(1).ControlCount=   2
      Begin VB.CommandButton cmdMyReq 
         Caption         =   "&View My Requisitions"
         Height          =   375
         Left            =   -74760
         TabIndex        =   39
         Top             =   7320
         Width           =   2415
      End
      Begin VB.Frame Frame1 
         Height          =   7575
         Left            =   120
         TabIndex        =   15
         Top             =   80
         Width           =   10695
         Begin VB.TextBox txtPss 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5040
            TabIndex        =   6
            Top             =   1800
            Width           =   615
         End
         Begin VB.TextBox txtLDtls 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2040
            TabIndex        =   4
            Top             =   1440
            Width           =   3615
         End
         Begin VB.TextBox txtRemarks 
            Height          =   285
            Left            =   2040
            TabIndex        =   7
            Top             =   2160
            Width           =   3615
         End
         Begin VB.Frame Frame2 
            Height          =   855
            Left            =   5760
            TabIndex        =   16
            Top             =   2880
            Width           =   4815
            Begin VB.CommandButton cmdRequest 
               Caption         =   "&Request"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   120
               TabIndex        =   9
               Top             =   240
               Width           =   2295
            End
            Begin VB.CommandButton cmdCancel 
               Caption         =   "&Cancel"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   2520
               TabIndex        =   10
               Top             =   240
               Width           =   2175
            End
         End
         Begin MSDataListLib.DataCombo DcmbLocCity 
            Height          =   330
            Left            =   2040
            TabIndex        =   3
            Top             =   1080
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   582
            _Version        =   393216
            MatchEntry      =   -1  'True
            Text            =   ""
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
         Begin MSDataListLib.DataCombo DcmbComCode 
            Height          =   330
            Left            =   2040
            TabIndex        =   2
            Top             =   720
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   582
            _Version        =   393216
            Style           =   2
            Text            =   ""
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
         Begin MSDataListLib.DataList DataList1 
            Height          =   2205
            Left            =   5760
            TabIndex        =   17
            Top             =   600
            Width           =   4815
            _ExtentX        =   8493
            _ExtentY        =   3889
            _Version        =   393216
         End
         Begin MSComCtl2.DTPicker DTPDate 
            Height          =   375
            Left            =   2040
            TabIndex        =   0
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   106430465
            CurrentDate     =   43294
         End
         Begin MSDataListLib.DataCombo DcmbVCat 
            Height          =   315
            Left            =   2040
            TabIndex        =   8
            Top             =   2520
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSComCtl2.DTPicker DTPTime 
            Height          =   375
            Left            =   4080
            TabIndex        =   1
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   106430466
            CurrentDate     =   43294
         End
         Begin MSFlexGridLib.MSFlexGrid MSFRequest 
            Height          =   3255
            Left            =   120
            TabIndex        =   35
            Top             =   3840
            Width           =   10455
            _ExtentX        =   18441
            _ExtentY        =   5741
            _Version        =   393216
            BackColor       =   16644579
         End
         Begin MSDataListLib.DataCombo DcmbRcat 
            Height          =   330
            Left            =   2040
            TabIndex        =   5
            Top             =   1800
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   582
            _Version        =   393216
            Style           =   2
            Text            =   ""
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
         Begin VB.Label Label11 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Passengers"
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
            Left            =   3840
            TabIndex        =   38
            Top             =   1800
            Width           =   1095
         End
         Begin VB.Label Label10 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Request Category"
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
            TabIndex        =   37
            Top             =   1800
            Width           =   1815
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Vehicle Category"
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
            TabIndex        =   34
            Top             =   2520
            Width           =   1815
         End
         Begin VB.Label Label2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Location (City)"
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
            TabIndex        =   33
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Label Label3 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Location Details"
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
            TabIndex        =   32
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label Label4 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Required Date - Time"
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
            TabIndex        =   31
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label5 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Remarks"
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
            TabIndex        =   30
            Top             =   2160
            Width           =   1815
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Caption         =   "Frequently Travelled List"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5760
            TabIndex        =   29
            Top             =   240
            Width           =   4815
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Caption         =   "Approx. (Km)"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   2880
            Width           =   1335
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Caption         =   "Approx. (Rs.)"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1560
            TabIndex        =   27
            Top             =   2880
            Width           =   1575
         End
         Begin VB.Label Label9 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Facilitator"
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
            TabIndex        =   26
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label lblDistance 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calisto MT"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   615
            Left            =   120
            TabIndex        =   25
            Top             =   3120
            Width           =   1335
         End
         Begin VB.Label lblValue 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calisto MT"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   615
            Left            =   1560
            TabIndex        =   24
            Top             =   3120
            Width           =   1575
         End
         Begin VB.Label Label12 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Reporting Person eMail"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   7200
            Width           =   1815
         End
         Begin VB.Label lblRPerson 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2040
            TabIndex        =   22
            Top             =   7200
            Width           =   3135
         End
         Begin VB.Label Label14 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Reporting Person Contact No."
            Height          =   255
            Left            =   5400
            TabIndex        =   21
            Top             =   7200
            Width           =   2175
         End
         Begin VB.Label lblRCont1 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   7680
            TabIndex        =   20
            Top             =   7200
            Width           =   2895
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Caption         =   "Balance Budget Amount. (Rs.)"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3240
            TabIndex        =   19
            Top             =   2880
            Width           =   2415
         End
         Begin VB.Label lblBudget 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calisto MT"
               Size            =   20.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   3240
            TabIndex        =   18
            Top             =   3120
            Width           =   2415
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSFMYRequest 
         Height          =   7095
         Left            =   -74880
         TabIndex        =   36
         Top             =   120
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   12515
         _Version        =   393216
         BackColor       =   16777215
      End
   End
   Begin VB.Frame Frame3 
      Height          =   8295
      Left            =   11160
      TabIndex        =   13
      Top             =   720
      Width           =   975
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
         Height          =   735
         Left            =   120
         Picture         =   "frmVReq.frx":0038
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Exit from the System"
         Top             =   7440
         Width           =   735
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   9075
      Width           =   12210
      _ExtentX        =   21537
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Bevel           =   2
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "8/17/2020"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Bevel           =   2
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "9:06 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Bevel           =   2
            Enabled         =   0   'False
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Bevel           =   2
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   15875
            MinWidth        =   15875
         EndProperty
      EndProperty
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   2400
      Picture         =   "frmVReq.frx":087D
      Stretch         =   -1  'True
      Top             =   120
      Width           =   6255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   14  'Copy Pen
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   2280
      Top             =   0
      Width           =   11175
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   0
      Picture         =   "frmVReq.frx":415A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "frmVReq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim T_Value As Double
Dim T_Distance As Long
Dim F_Loaded As Integer

Private Sub cmdCancel_Click()
    On Error GoTo er_EH:
    If MsgBox("Are you sure Do you Want to Delete your selected VR ?", vbQuestion + vbYesNo, "Delete Record") = vbYes Then
        PR_HRS_Open_CON
        HRS.Execute "Delete from HRS_TR_Request Where VR_NO=" & Val(MSFRequest.TextMatrix(MSFRequest.Row, 0))
        PR_HRS_Close_CON
        PR_Grid
        PR_Grid_My_Request
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

Private Sub cmdMyReq_Click()
    PR_Grid_My_Request
End Sub

Private Sub cmdRequest_Click()
    On Error GoTo er_EH:
    Dim R_ComCode, R_Loc_Dtls, R_Remarks, R_Date_Time, R_Method As String
    Dim Pss As Integer
    
    If Trim(DcmbComCode.Text) = "" Or Trim(DcmbComCode.Text) = "-Select-" Then
        MsgBox "Starting Place NOT Selected", vbExclamation
        DcmbComCode.SetFocus
        Exit Sub
    End If
    
    If Trim(DcmbLocCity.Text) = "" Or Trim(DcmbLocCity.Text) = "-Select-" Then
        MsgBox "Travelling City NOT Selected", vbExclamation
        DcmbLocCity.SetFocus
        Exit Sub
    End If
    
    If Trim(txtLDtls.Text) = "" Then
        MsgBox "Travelling Location Details NOT Entered", vbExclamation
        txtLDtls.SetFocus
        Exit Sub
    End If
    
    If Trim(DcmbVCat.Text) = "" Or Trim(DcmbVCat.Text) = "-Select-" Then
        MsgBox "Vehicle Category NOT Selected", vbExclamation
        DcmbVCat.SetFocus
        Exit Sub
    End If
    
    If Trim(txtPss.Text) = "" Then
        MsgBox "Number Of Passengers NOT Entered", vbExclamation
        txtPss.SetFocus
        Exit Sub
    End If
    
    R_ComCode = Trim(DcmbComCode.Text)
    R_Loc_Dtls = Trim(txtLDtls.Text)
    R_Remarks = Trim(txtRemarks.Text)
    R_Date_Time = Format(DTPDate.Value, "MM/dd/yyyy") & " " & Format(DTPTime.Value, "hh:mm:ss")
    R_Method = Trim(DcmbRcat.Text)
    Pss = Val(txtPss.Text)
    FN_Find_VCat_ID Trim(DcmbVCat.Text)
    FN_Find_Reason_Cat_ID Trim(DcmbRcat.Text), Trim(DcmbComCode.Text)
    
    PR_HRS_Open_CON
    HRS.Execute "INSERT INTO HRS_TR_Request(Req_Date_Time,Com_Code,Dist_ID,Loc_Dtls,Remarks,Cat_ID,U_ID,Status_ID,T_Date_Time,Reason_Cat_ID,Req_Passengers)" _
            + "Values('" & R_Date_Time & "','" & R_ComCode & "'," & Loc_City_ID & ",'" & R_Loc_Dtls & "','" & R_Remarks & "'," & V_Cat_ID & ",'" & UserID & "',1,'" & Format(Date, "MM/dd/yyyy") + " " + Format(Time, "hh:mm:ss") & "'," & R_Method_ID & "," & Pss & ")"
    PR_HRS_Close_CON

    Set RS = New ADODB.Recordset
    Dim Req_No As Long
    PR_HRS_Open_CON
    RS.Open "Select isnull(max(VR_NO),0) as Max_No from HRS_TR_Request Where U_ID='" & UserID & "'", HRS, adOpenStatic, adLockReadOnly
    If RS.EOF = False Then
        Req_No = Val(RS!Max_No)
    Else
        Req_No = 0
    End If
    SMS_Message = "Please Approve My VR#" & Str(Req_No) & " " & " To " & R_Loc_Dtls & " Balance Budget is Rs." & Str(Balance_Budget)
    
    PR_HRS_Close_CON
    FN_Find_Reason_Cat_ID Trim(DcmbRcat.Text), Trim(DcmbComCode.Text)
        If Format(Time, "HH:mm") >= CDate("16:30") And DTPDate.Value = Date + 1 Then
        MsgBox "Your Vehicle Requisition Should have to Approved by Finance Division", vbExclamation
        FN_Status_Update 8, Req_No, Module_ID, Sub_Module_ID
    Else
        If Balance_Budget - Val(lblValue.Caption) < 0 Then
            MsgBox "Your Vehicle Requisition Should have to Approved by Finance Division", vbExclamation
            FN_Status_Update 8, Req_No, Module_ID, Sub_Module_ID
        Else
            FN_Status_Update 1, Req_No, Module_ID, Sub_Module_ID
        End If
    End If
    
    PR_Grid
    PR_Grid_My_Request
    
    PR_HRS_Open_CON
    Set RS = New ADODB.Recordset
    RS.Open "SELECT * FROM HRSV_sys_MSTR_Employee Where U_ID='" & UserID & "'", HRS, adOpenStatic, adLockReadOnly
    
    Mod_SMS.FN_SMS RS!R_Contact1, SMS_Message
    PR_HRS_Close_CON
    Exit Sub
    
er_EH:
    MsgBox Err.Description, vbCritical
    Err.Clear
End Sub

Private Sub DcmbComCode_Change()
    If Trim(DcmbComCode.Text) <> "" And Trim(DcmbComCode.Text) <> "-Select-" Then
        PR_Fill_City
        PR_Fill_Reason_Category
        If Trim(DcmbLocCity.Text) = "" And Trim(DcmbLocCity.Text) = "-Select-" Then
            PR_Cost_Calc
        End If
    End If
End Sub

Private Sub DcmbComCode_Click(Area As Integer)
    If Trim(DcmbLocCity.Text) <> "" And Trim(DcmbLocCity.Text) <> "-Select-" Then
        Dim From_Loc, To_Loc As String
        From_Loc = Trim(DcmbComCode.Text)
        To_Loc = Trim(DcmbLocCity.Text)
        PR_Cost_Calc
    Else
        Exit Sub
    End If
End Sub

Private Sub DcmbLocCity_Click(Area As Integer)
    If Trim(DcmbLocCity.Text) <> "" And Trim(DcmbLocCity.Text) <> "-Select-" Then
        PR_Cost_Calc
    End If
End Sub

Private Sub DcmbVCat_Change()
    PR_Cost_Calc
End Sub

Private Sub Form_Load()
    DTPDate.Value = Date
    PR_Fill_Combo
    PR_Reporting
    PR_Grid
    Module_ID = 1
    Sub_Module_ID = 1
    PR_Budget_Validation
End Sub

Public Sub PR_Fill_Combo()
    PR_HRS_Open_CON
    Set RS = New ADODB.Recordset
    RS.Open "Select Com_Code from HRS_sys_Company Order by Com_Code", HRS, adOpenKeyset, adLockReadOnly
    DcmbComCode.ListField = "Com_Code"
    Set DcmbComCode.RowSource = RS
    
    Set RS = New ADODB.Recordset
    RS.Open "SELECT Category FROM HRS_TR_MSTR_Category Order by Category", HRS, adOpenKeyset, adLockReadOnly
    DcmbVCat.ListField = "Category"
    Set DcmbVCat.RowSource = RS
    DcmbVCat.Text = "-Select-"
End Sub

Public Sub PR_Fill_City()
    PR_HRS_Open_CON
    Set RS = New ADODB.Recordset
    RS.Open "Select Distinct To_City from HRS_TR_MSTR_Distance Where From_Loc='" & Trim(DcmbComCode.Text) & "'", HRS, adOpenKeyset, adLockReadOnly
    DcmbLocCity.ListField = "To_City"
    Set DcmbLocCity.RowSource = RS
    DcmbLocCity.Text = "-Select-"
End Sub

Public Sub PR_Reporting()
    Set RS = New ADODB.Recordset
    PR_HRS_Open_CON
    RS.Open "Select * from HRSV_sys_MSTR_Employee Where U_ID='" & UserID & "'", HRS, adOpenStatic, adLockReadOnly
    If RS.EOF = False Then
        lblRPerson.Caption = Trim(RS!R_Email)
        lblRCont1.Caption = Trim(RS!R_Contact1)
    Else
        lblRPerson.Caption = "-"
        lblRCont1.Caption = "-"
    End If
    PR_HRS_Close_CON
End Sub

Public Sub PR_Cost_Calc()
    FN_Find_Distance Trim(DcmbComCode.Text), Trim(DcmbLocCity.Text)
    If Trim(DcmbVCat.Text) <> "" And Trim(DcmbVCat.Text) <> "-Select-" Then
        V_Category = Trim(DcmbVCat.Text)
        FN_Find_VCat_ID V_Category
        lblDistance.Caption = Distance
        lblValue.Caption = Distance * Rate
        lblBudget.Caption = Balance_Budget - (Distance * Rate)
    Else
        lblDistance.Caption = ""
        lblValue.Caption = ""
        Exit Sub
    End If
End Sub

Public Sub PR_Grid()
    MSFRequest.Cols = 7
    R = 1
    Set GRDRS = New ADODB.Recordset
    PR_HRS_Open_CON
    GRDRS.Open "SELECT * from HRSV_TR_Requests WHERE U_ID='" & UserID & "' and Status_ID IN(1,8)", HRS, adOpenStatic, adLockReadOnly
    If GRDRS.EOF = False Then
        Do While GRDRS.EOF = False
            MSFRequest.Rows = R + 1
            MSFRequest.TextMatrix(R, 0) = Format(Val(GRDRS!VR_No), "0000#")
            MSFRequest.TextMatrix(R, 1) = Trim(GRDRS!Req_ComCode)
            MSFRequest.TextMatrix(R, 2) = Trim(GRDRS!To_City)
            MSFRequest.TextMatrix(R, 3) = Trim(GRDRS!To_Loc)
            If IsNull(GRDRS!Remarks) = False Then
                MSFRequest.TextMatrix(R, 4) = Trim(GRDRS!Remarks)
            Else
                MSFRequest.TextMatrix(R, 4) = ""
            End If
            MSFRequest.TextMatrix(R, 5) = Format(Trim(GRDRS!Req_Date_Time), "dd-MMM-YYYY")
            MSFRequest.TextMatrix(R, 6) = Format(Trim(GRDRS!Req_Date_Time), "HH:mm")
            R = R + 1
            GRDRS.MoveNext
        Loop
    Else
        MSFRequest.Rows = 1
    End If
    
    MSFRequest.TextMatrix(0, 0) = "REQ. #"
    MSFRequest.TextMatrix(0, 1) = "FROM LOCTAION"
    MSFRequest.TextMatrix(0, 2) = "TO CITY"
    MSFRequest.TextMatrix(0, 3) = "LOCATION DETAILS"
    MSFRequest.TextMatrix(0, 4) = "REMARKS"
    MSFRequest.TextMatrix(0, 5) = "REQ. DATE"
    MSFRequest.TextMatrix(0, 6) = "REQ. TIME"
    MSFRequest.ColWidth(0) = 1000
    MSFRequest.ColWidth(1) = 1200
    MSFRequest.ColWidth(2) = 2000
    MSFRequest.ColWidth(3) = 2500
    MSFRequest.ColWidth(4) = 1500
    MSFRequest.ColWidth(5) = 1050
    MSFRequest.ColWidth(5) = 1050
    
    PR_HRS_Close_CON
End Sub

Public Sub PR_Grid_My_Request()
    MSFMYRequest.Cols = 7
    R = 1
    Set GRDRS = New ADODB.Recordset
    PR_HRS_Open_CON
    GRDRS.Open "SELECT * FROM HRSV_TR_Requests WHERE U_ID='" & UserID & "' and Status_ID <>6", HRS, adOpenStatic, adLockReadOnly
    If GRDRS.EOF = False Then
        Do While GRDRS.EOF = False
            MSFMYRequest.Rows = R + 1
            MSFMYRequest.TextMatrix(R, 0) = Format(Val(GRDRS!VR_No), "0000#")
            MSFMYRequest.TextMatrix(R, 1) = Trim(GRDRS!Req_ComCode)
            MSFMYRequest.TextMatrix(R, 2) = Trim(GRDRS!To_City)
            MSFMYRequest.TextMatrix(R, 3) = Trim(GRDRS!To_Loc)
            MSFMYRequest.TextMatrix(R, 4) = Trim(GRDRS!Remarks)
            MSFMYRequest.TextMatrix(R, 5) = Format(Trim(GRDRS!Req_Date_Time), "dd-MMM-YYYY")
            MSFMYRequest.TextMatrix(R, 6) = Format(Trim(GRDRS!Req_Date_Time), "HH:mm")
            R = R + 1
            GRDRS.MoveNext
        Loop
    Else
        MSFMYRequest.Rows = 1
    End If
    
    MSFMYRequest.TextMatrix(0, 0) = "REQ. #"
    MSFMYRequest.TextMatrix(0, 1) = "FROM LOCTAION"
    MSFMYRequest.TextMatrix(0, 2) = "TO CITY"
    MSFMYRequest.TextMatrix(0, 3) = "LOCATION DETAILS"
    MSFMYRequest.TextMatrix(0, 4) = "REMARKS"
    MSFMYRequest.TextMatrix(0, 5) = "REQ. DATE"
    MSFMYRequest.TextMatrix(0, 6) = "REQ. TIME"
    MSFMYRequest.ColWidth(0) = 1000
    MSFMYRequest.ColWidth(1) = 1200
    MSFMYRequest.ColWidth(2) = 2000
    MSFMYRequest.ColWidth(3) = 2500
    MSFMYRequest.ColWidth(4) = 1500
    MSFMYRequest.ColWidth(5) = 1050
    MSFMYRequest.ColWidth(5) = 1050
    

    For R = 1 To MSFMYRequest.Rows - 1
    Set GRDRS = New ADODB.Recordset
    GRDRS.Open "Select isnull(Status_ID,0) as Status_ID from HRSV_TR_Requests Where VR_No=" & Val(MSFMYRequest.TextMatrix(R, 0)), HRS, adOpenStatic, adLockReadOnly
        For C = 1 To MSFMYRequest.Cols - 1
            MSFMYRequest.Col = C
            MSFMYRequest.Row = R
            If GRDRS.EOF = False Then
                If Val(GRDRS!Status_ID) = 3 Then
                    MSFMYRequest.CellBackColor = &HC0C0FF
                End If
                
                If Val(GRDRS!Status_ID) = 2 Then
                    MSFMYRequest.CellBackColor = &HC0FFC0
                End If
                
                If Val(GRDRS!Status_ID) = 5 Then
                    MSFMYRequest.CellBackColor = &HFF8080
                End If
            End If
        Next C
    Next R
    PR_HRS_Close_CON
End Sub

Private Sub lblBudget_Change()
    If Val(lblBudget.Caption) < 0 Then
        lblBudget.ForeColor = &H80&
    Else
        lblBudget.ForeColor = &H80000012
    End If
End Sub

Private Sub MSFMYRequest_DblClick()
    If MSFMYRequest.Rows > 1 Then
        Tmp = Trim(MSFMYRequest.TextMatrix(MSFMYRequest.Row, 0))
        Load frmEscalation
        frmEscalation.Show (1)
    End If
End Sub

Public Sub PR_Fill_Reason_Category()
    PR_HRS_Open_CON
    Set RS = New ADODB.Recordset
    RS.Open "SELECT Distinct Reason_Category FROM HRS_TR_MSTR_Reason_Cat Where Module_ID=1 and Sub_Module_ID=1 AND Com_Code='" & Trim(DcmbComCode.Text) & "' Order by Reason_Category", HRS, adOpenKeyset, adLockReadOnly
    DcmbRcat.ListField = "Reason_Category"
    Set DcmbRcat.RowSource = RS
    DcmbRcat.Text = "Adhoc"
    PR_HRS_Close_CON
End Sub
Public Sub PR_Budget_Validation()
    PR_HRS_Open_CON
    Set RS = New ADODB.Recordset
    RS.Open "Select * from HRSV_TR_RPT_Division_Costs_Balance Where D_Code=" & U_D_Code & " and Month=" & Month(Date) & " and Year=" & Year(Date), HRS, adOpenStatic, adLockReadOnly
    If RS.EOF = False Then
        Balance_Budget = Val(RS!Balance)
    Else
        Balance_Budget = 0
    End If
    lblBudget.Caption = Balance_Budget
    PR_HRS_Close_CON
End Sub

