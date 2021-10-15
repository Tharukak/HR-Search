VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmVR 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Vehicle Requisition"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13380
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
   ScaleHeight     =   8250
   ScaleWidth      =   13380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   6255
      Left            =   6000
      TabIndex        =   4
      Top             =   1680
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   11033
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   7995
      Width           =   13380
      _ExtentX        =   23601
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
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
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6255
      Left            =   0
      TabIndex        =   2
      Top             =   1680
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   11033
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Request Vehicle"
      TabPicture(0)   =   "frmVR.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Request Adjustments"
      TabPicture(1)   =   "frmVR.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      Begin VB.Frame Frame1 
         Caption         =   "Current Day Requisitions"
         Height          =   5775
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   5655
         Begin VB.CommandButton cmdCancel 
            Height          =   615
            Left            =   4920
            Picture         =   "frmVR.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   3360
            Width           =   615
         End
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
            Height          =   1575
            Left            =   120
            TabIndex        =   24
            Top             =   4080
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   2778
            _Version        =   393216
         End
         Begin VB.CommandButton cmdRequest 
            Height          =   615
            Left            =   3240
            Picture         =   "frmVR.frx":053E
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   3360
            Width           =   1695
         End
         Begin MSComCtl2.DTPicker DTPRTime 
            Height          =   300
            Left            =   1440
            TabIndex        =   22
            Top             =   3720
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "hh:mm"
            Format          =   108855298
            CurrentDate     =   43262
         End
         Begin MSComCtl2.DTPicker DTPRDate 
            Height          =   300
            Left            =   1440
            TabIndex        =   21
            Top             =   3360
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   108855299
            CurrentDate     =   43262
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   1440
            TabIndex        =   20
            Top             =   3000
            Width           =   4095
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   2760
            TabIndex        =   17
            Top             =   1680
            Width           =   2775
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   120
            TabIndex        =   16
            Top             =   1680
            Width           =   2535
         End
         Begin VB.ListBox List2 
            Height          =   840
            Left            =   2760
            TabIndex        =   15
            Top             =   2040
            Width           =   2775
         End
         Begin VB.ListBox List1 
            Height          =   840
            Left            =   120
            TabIndex        =   14
            Top             =   2040
            Width           =   2535
         End
         Begin VB.ComboBox cmbComCode 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   240
            Width           =   1815
         End
         Begin VB.ComboBox cmbVehicle 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "- To -"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   255
            Left            =   2760
            TabIndex        =   19
            Top             =   1320
            Width           =   2775
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "- From -"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   1320
            Width           =   2535
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Request From"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label13 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Time"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   3720
            Width           =   1215
         End
         Begin VB.Label Label12 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Date"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   3360
            Width           =   1215
         End
         Begin VB.Label Label10 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Location Details"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   3000
            Width           =   1215
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            Caption         =   "Location"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   960
            Width           =   5415
         End
         Begin VB.Label Label8 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Vehicle Type"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   600
            Width           =   1215
         End
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   0
      X2              =   13320
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Reporting e-Mail"
      Height          =   255
      Left            =   5880
      TabIndex        =   31
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Division Name"
      Height          =   255
      Left            =   5880
      TabIndex        =   30
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Reporting to"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "User Name"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblREmail 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7320
      TabIndex        =   27
      Top             =   1200
      Width           =   6015
   End
   Begin VB.Label lblDivision 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7320
      TabIndex        =   26
      Top             =   840
      Width           =   6015
   End
   Begin VB.Label lblRPerson 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   1200
      Width           =   4455
   End
   Begin VB.Label lblUserName 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   0
      Top             =   840
      Width           =   4455
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   2280
      Picture         =   "frmVR.frx":0F72
      Stretch         =   -1  'True
      Top             =   240
      Width           =   4695
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
      Picture         =   "frmVR.frx":299F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "frmVR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Set RS = New ADODB.Recordset
    PR_HRS_Open_CON
    cmbComCode.Clear
    RS.Open "Select Com_Code from HRS_sys_Company Order by Com_Code", HRS, adOpenStatic, adLockReadOnly
    Do While RS.EOF = False
        cmbComCode.AddItem Trim(RS!Com_Code)
        RS.MoveNext
    Loop
    
    Set RS = New ADODB.Recordset
    cmbVehicle.Clear
    RS.Open "Select Vehicle_type from HRS_sys_Vehicles  Order by Vehicle_ID", HRS, adOpenStatic, adLockReadOnly
    Do While RS.EOF = False
        cmbVehicle.AddItem Trim(RS!Vehicle_type)
        RS.MoveNext
    Loop
    
    Set RS = New ADODB.Recordset
    RS.Open "Select U_FName,R_Name,D_Name,R_EMail from HRSV_sys_Employee Where U_ID=" & UserID, HRS, adOpenStatic, adLockReadOnly
    If RS.EOF = False Then
        lblUserName.Caption = Trim(RS!U_Name)
        lblRPerson.Caption = Trim(RS!R_Name)
        lblDivision.Caption = Trim(RS!D_Name)
        lblREmail.Caption = Trim(RS!R_Email)
    Else
        lblUserName.Caption = ""
        lblRPerson.Caption = ""
        lblDivision.Caption = ""
        lblREmail.Caption = ""
    End If
    PR_HRS_Close_CON
End Sub
