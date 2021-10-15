VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmSpash 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Welcome to HR Search - Exteme"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5955
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   120
      Top             =   960
   End
   Begin MSComctlLib.StatusBar Stat 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   3555
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   12347
            MinWidth        =   12347
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   3240
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   0
      Picture         =   "frmSpash.frx":0000
      ScaleHeight     =   3225
      ScaleWidth      =   5985
      TabIndex        =   2
      Top             =   0
      Width           =   6015
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Brandix Apparel Solutions - Lingerie"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   5175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "HR Search - Extream"
         BeginProperty Font 
            Name            =   "Neuropolitical Rg"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   0
         Width           =   3375
      End
   End
End
Attribute VB_Name = "frmSpash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Stat.Panels(1).Text = "Version " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub timer1_Timer()
    PB.Max = 100
    If PB.Value < 100 Then
        PB.Value = PB.Value + 1
    Else
        Load frmLogin
        frmLogin.Show
        Unload Me
    End If
End Sub
