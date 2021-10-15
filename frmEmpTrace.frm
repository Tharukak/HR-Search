VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEmpTrace 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Associate Traceability"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15705
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
   ScaleHeight     =   8730
   ScaleWidth      =   15705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   7200
      Left            =   14880
      TabIndex        =   26
      Top             =   1200
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
         Picture         =   "frmEmpTrace.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Exit from the System"
         Top             =   6600
         Width           =   495
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7095
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   12515
      _Version        =   393216
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   14737632
      TabCaption(0)   =   "01st Closed Ones"
      TabPicture(0)   =   "frmEmpTrace.frx":0845
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Text1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "MSFlexGrid1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Text3"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "02nd Closed Ones"
      TabPicture(1)   =   "frmEmpTrace.frx":0861
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text6"
      Tab(1).Control(1)=   "Text5"
      Tab(1).Control(2)=   "Text4"
      Tab(1).Control(3)=   "MSFlexGrid2"
      Tab(1).Control(4)=   "Label7"
      Tab(1).Control(5)=   "Label6"
      Tab(1).Control(6)=   "Label5"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "03rd Closed Ones"
      TabPicture(2)   =   "frmEmpTrace.frx":087D
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Text9"
      Tab(2).Control(1)=   "Text8"
      Tab(2).Control(2)=   "Text7"
      Tab(2).Control(3)=   "MSFlexGrid3"
      Tab(2).Control(4)=   "Label10"
      Tab(2).Control(5)=   "Label9"
      Tab(2).Control(6)=   "Label8"
      Tab(2).ControlCount=   7
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   -73200
         TabIndex        =   25
         Top             =   1200
         Width           =   8655
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   -73200
         TabIndex        =   24
         Top             =   840
         Width           =   2655
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   -73200
         TabIndex        =   21
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   -73200
         TabIndex        =   19
         Top             =   1200
         Width           =   8655
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   -73200
         TabIndex        =   18
         Top             =   840
         Width           =   2655
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   -73200
         TabIndex        =   15
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1800
         TabIndex        =   13
         Top             =   1320
         Width           =   8655
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1800
         TabIndex        =   12
         Top             =   960
         Width           =   2655
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   5175
         Left            =   240
         TabIndex        =   7
         Top             =   1680
         Width           =   14295
         _ExtentX        =   25215
         _ExtentY        =   9128
         _Version        =   393216
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1800
         TabIndex        =   6
         Top             =   600
         Width           =   2655
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   5175
         Left            =   -74880
         TabIndex        =   8
         Top             =   1680
         Width           =   14295
         _ExtentX        =   25215
         _ExtentY        =   9128
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
         Height          =   5415
         Left            =   -74880
         TabIndex        =   9
         Top             =   1560
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   9551
         _Version        =   393216
      End
      Begin VB.Label Label10 
         Caption         =   "Contact Number"
         Height          =   255
         Left            =   -74760
         TabIndex        =   23
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Name"
         Height          =   255
         Left            =   -74760
         TabIndex        =   22
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "NIC Number"
         Height          =   255
         Left            =   -74760
         TabIndex        =   20
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Contact Number"
         Height          =   255
         Left            =   -74760
         TabIndex        =   17
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Name"
         Height          =   255
         Left            =   -74760
         TabIndex        =   16
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "NIC Number"
         Height          =   255
         Left            =   -74760
         TabIndex        =   14
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Contact Number"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Name"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "NIC Number"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.TextBox txtEmpNo 
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Top             =   840
      Width           =   2055
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   8475
      Width           =   15705
      _ExtentX        =   27702
      _ExtentY        =   450
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
   Begin VB.Label Label11 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Employee Number"
      Height          =   255
      Left            =   3840
      TabIndex        =   28
      Top             =   840
      Width           =   11775
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   0
      Picture         =   "frmEmpTrace.frx":0899
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Associate Traceability"
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
      TabIndex        =   4
      Top             =   120
      Width           =   5655
   End
   Begin VB.Label Label1 
      Caption         =   "Employee Number"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1455
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
End
Attribute VB_Name = "frmEmpTrace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    Close All
    Unload Me
End Sub
