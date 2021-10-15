VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.MDIForm MDIHome 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Welcome to HR Search"
   ClientHeight    =   3015
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   4560
   Icon            =   "MDIHome.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIHome.frx":4B85A
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   10260
      Width           =   20250
      _ExtentX        =   35719
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   14111
            MinWidth        =   14111
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "7/8/2021"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "2:15 PM"
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   46
      ImageHeight     =   46
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIHome.frx":94CA5
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIHome.frx":9565C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIHome.frx":95ECB
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIHome.frx":96C91
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIHome.frx":97817
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIHome.frx":984C0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20250
      _ExtentX        =   35719
      _ExtentY        =   1429
      ButtonWidth     =   4445
      ButtonHeight    =   1376
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Vehicle Requisition"
            Key             =   "VR"
            Object.ToolTipText     =   "Vehicle Requisition"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Approval Pannel"
            Key             =   "Approval"
            Object.ToolTipText     =   "Approval Pannel"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Vehicle Arrangement"
            Key             =   "VA"
            Object.ToolTipText     =   "Vehicle Arrangement"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Meter Readings"
            Key             =   "Settings"
            Object.ToolTipText     =   "Meter Reading"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Meal"
            Key             =   "Meal"
            Object.ToolTipText     =   "Meal Module"
            Object.Tag             =   "Meal"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reports"
            Key             =   "Reports"
            Object.ToolTipText     =   "Reports"
            Object.Tag             =   "Reports"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu SM1 
         Caption         =   "&Adhoc Transport Module"
         Begin VB.Menu SM11 
            Caption         =   "&Vehicle Request Form"
         End
         Begin VB.Menu SM12 
            Caption         =   "&Approval Pannel"
         End
         Begin VB.Menu SM13 
            Caption         =   "&Vehicle Arrangement"
         End
         Begin VB.Menu SM14 
            Caption         =   "&Meter Reading"
         End
      End
      Begin VB.Menu mnuMM2 
         Caption         =   "&Fixed Transport Module"
         Begin VB.Menu SM21 
            Caption         =   "&Root Allocatnion"
         End
      End
      Begin VB.Menu mnuMM3 
         Caption         =   "&Meal Count"
         Begin VB.Menu SM31 
            Caption         =   "Meal Count Dashboard"
         End
         Begin VB.Menu SM32 
            Caption         =   "&CSV Generator"
         End
      End
      Begin VB.Menu mnuMM4 
         Caption         =   "&Barcode Creator"
         Begin VB.Menu SM41 
            Caption         =   "&Employee Barcode Creator"
         End
      End
      Begin VB.Menu MnuMM5 
         Caption         =   "&Health Care"
         Begin VB.Menu SM51 
            Caption         =   "&Medical Register"
         End
      End
      Begin VB.Menu MnuMM6 
         Caption         =   "&Human Resource Management"
         Begin VB.Menu MnuMM61 
            Caption         =   "&Employee Master File"
         End
         Begin VB.Menu MnuMM62 
            Caption         =   "&Associate Treceability"
         End
      End
      Begin VB.Menu MM6 
         Caption         =   "&Reports"
      End
      Begin VB.Menu MB1 
         Caption         =   "-"
      End
      Begin VB.Menu MM7 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuConfig 
      Caption         =   "&Confirguarations"
      Begin VB.Menu Con_MM1 
         Caption         =   "&Vehicle Master File"
      End
      Begin VB.Menu Con_MM2 
         Caption         =   "&Employee Masterfile"
      End
   End
   Begin VB.Menu MnuSecurity 
      Caption         =   "&Security"
      Begin VB.Menu Sec_MM1 
         Caption         =   "&User Rights"
      End
   End
   Begin VB.Menu MnuUtilities 
      Caption         =   "&Utilities"
      Begin VB.Menu Util_MM1 
         Caption         =   "&Change Password"
      End
   End
End
Attribute VB_Name = "MDIHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Con_MM1_Click()
    MenuAccess 7
    If RightsMode = 1 Then
        Load frmVMF
        frmVMF.Show (1)
    End If
End Sub

Private Sub Con_MM2_Click()
    MenuAccess 2
    If RightsMode = 1 Then
        Load frmEMF
        frmEMF.Show (1)
    End If
End Sub

Private Sub MM6_Click()
    Load frmReports
    frmReports.Show (1)
End Sub

Private Sub MM7_Click()
    Close All
    Unload Me
End Sub

Private Sub MnuMM61_Click()
    Load frmExtEmpMSTR
    frmExtEmpMSTR.Show (1)
End Sub

Private Sub MnuMM62_Click()
    Load frmEmpTrace
    frmEmpTrace.Show (1)
End Sub

Private Sub Sec_MM1_Click()
    MenuAccess 1
    If RightsMode = 1 Then
        Load frmSecurity
        frmSecurity.Show (1)
    End If
End Sub

Private Sub SM11_Click()
    MenuAccess 3
    If RightsMode = 1 Then
        Load frmVReq
        frmVReq.Show (1)
    End If
End Sub

Private Sub SM12_Click()
    MenuAccess 4
    If RightsMode = 1 Then
        Load frmApproval
        frmApproval.Show (1)
    End If
End Sub

Private Sub SM13_Click()
    MenuAccess 7
    If RightsMode = 1 Then
        Load frmVArrange
        frmVArrange.Show (1)
    End If
End Sub

Private Sub SM14_Click()
    MenuAccess 8
    If RightsMode = 1 Then
        Load frmMeter
        frmMeter.Show (1)
    End If
End Sub

Private Sub SM21_Click()
    Load frmRootAllocation
    frmRootAllocation.Show (1)
End Sub

Private Sub SM31_Click()
    MenuAccess 14
    If RightsMode = 1 Then
        Load frmMealCount
        frmMealCount.Show (1)
    End If
End Sub

Private Sub SM32_Click()
    Load frmMealCsv
    frmMealCsv.Show (1)
End Sub

Private Sub SM41_Click()
    MenuAccess 13
    If RightsMode = 1 Then
        Load frmEmpBarcode
        frmEmpBarcode.Show (1)
    End If
End Sub

Private Sub SM51_Click()
    Load frmMedReg
    frmMedReg.Show (1)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "VR"
        SM11_Click
        Case "Approval"
        SM12_Click
        Case "VA"
        SM13_Click
        Case "Settings"
        SM14_Click
        Case "Reports"
        MM6_Click
        Case "Meal"
        SM31_Click
    End Select
End Sub

Private Sub Util_MM1_Click()
    Load frmChangePW
    frmChangePW.Show (1)
End Sub
