VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmExtEmpMSTR 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Extended Employee Master File"
   ClientHeight    =   10155
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14880
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
   ScaleHeight     =   10155
   ScaleWidth      =   14880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   7200
      TabIndex        =   41
      Top             =   3360
      Width           =   2775
   End
   Begin VB.TextBox txtContact1 
      Height          =   285
      Left            =   2280
      TabIndex        =   39
      Text            =   "`"
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Frame Frame2 
      Height          =   6120
      Left            =   14040
      TabIndex        =   25
      Top             =   3720
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
         Picture         =   "frmExtEmpMSTR.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Exit from the System"
         Top             =   5520
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Living Details"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   120
      TabIndex        =   24
      Top             =   3720
      Width           =   13815
      Begin VB.CommandButton cmdUpdate 
         Enabled         =   0   'False
         Height          =   975
         Left            =   12600
         Picture         =   "frmExtEmpMSTR.frx":0845
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   4680
         Width           =   975
      End
      Begin VB.CommandButton cmdSave 
         Height          =   975
         Left            =   9120
         Picture         =   "frmExtEmpMSTR.frx":152F
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   4680
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Enabled         =   0   'False
         Height          =   975
         Left            =   11520
         Picture         =   "frmExtEmpMSTR.frx":1D7A
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   4680
         Width           =   975
      End
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   9960
         TabIndex        =   69
         Top             =   4200
         Width           =   3615
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   9960
         TabIndex        =   68
         Top             =   3840
         Width           =   3615
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   9960
         TabIndex        =   67
         Top             =   3480
         Width           =   3615
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   9960
         TabIndex        =   66
         Top             =   3120
         Width           =   3615
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   6120
         TabIndex        =   65
         Top             =   4200
         Width           =   3615
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   6120
         TabIndex        =   64
         Top             =   3840
         Width           =   3615
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   6120
         TabIndex        =   63
         Top             =   3480
         Width           =   3615
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   6120
         TabIndex        =   62
         Top             =   3120
         Width           =   3615
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   2280
         TabIndex        =   61
         Top             =   4200
         Width           =   3615
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   2280
         TabIndex        =   60
         Top             =   3840
         Width           =   3615
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2280
         TabIndex        =   59
         Top             =   3480
         Width           =   3615
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2280
         TabIndex        =   58
         Top             =   3120
         Width           =   3615
      End
      Begin MSDataListLib.DataCombo dcmbRCat1 
         Height          =   315
         Left            =   2280
         TabIndex        =   43
         Top             =   600
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcmbRCat2 
         Height          =   315
         Left            =   6120
         TabIndex        =   44
         Top             =   600
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcmbRCat3 
         Height          =   315
         Left            =   9960
         TabIndex        =   45
         Top             =   600
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcmbProvince1 
         Height          =   315
         Left            =   2280
         TabIndex        =   46
         Top             =   960
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcmbProvince2 
         Height          =   315
         Left            =   6120
         TabIndex        =   47
         Top             =   960
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcmbProvince3 
         Height          =   315
         Left            =   9960
         TabIndex        =   48
         Top             =   960
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcmbPolice1 
         Height          =   315
         Left            =   2280
         TabIndex        =   49
         Top             =   1680
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcmbPolice2 
         Height          =   315
         Left            =   6120
         TabIndex        =   50
         Top             =   1680
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcmbPolice3 
         Height          =   315
         Left            =   9960
         TabIndex        =   51
         Top             =   1680
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcmbMOH1 
         Height          =   315
         Left            =   2280
         TabIndex        =   52
         Top             =   2040
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcmbMOH2 
         Height          =   315
         Left            =   6120
         TabIndex        =   53
         Top             =   2040
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcmbMOH3 
         Height          =   315
         Left            =   9960
         TabIndex        =   54
         Top             =   2040
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DataCombo13 
         Height          =   315
         Left            =   2280
         TabIndex        =   55
         Top             =   2400
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DataCombo14 
         Height          =   315
         Left            =   6120
         TabIndex        =   56
         Top             =   2400
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DataCombo15 
         Height          =   315
         Left            =   9960
         TabIndex        =   57
         Top             =   2400
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcmbDistrict1 
         Height          =   315
         Left            =   2280
         TabIndex        =   77
         Top             =   1320
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcmbDistrict2 
         Height          =   315
         Left            =   6120
         TabIndex        =   78
         Top             =   1320
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcmbDistrict3 
         Height          =   315
         Left            =   9960
         TabIndex        =   79
         Top             =   1320
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label Label24 
         BackColor       =   &H00FFFFFF&
         Caption         =   "District"
         Height          =   255
         Left            =   120
         TabIndex        =   76
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Update"
         Height          =   255
         Left            =   12600
         TabIndex        =   75
         Top             =   5640
         Width           =   975
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Edit"
         Height          =   255
         Left            =   11520
         TabIndex        =   74
         Top             =   5640
         Width           =   975
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Save"
         Height          =   255
         Left            =   9120
         TabIndex        =   73
         Top             =   5640
         Width           =   1095
      End
      Begin VB.Image Image3 
         BorderStyle     =   1  'Fixed Single
         Height          =   1455
         Left            =   9000
         Stretch         =   -1  'True
         Top             =   4560
         Width           =   4695
      End
      Begin VB.Label Label29 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Remarks"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   4200
         Width           =   1935
      End
      Begin VB.Label Label27 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nearest Hospital"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label Label26 
         BackColor       =   &H00FFFFFF&
         Caption         =   "MOH Office"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label25 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Police Area"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label23 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Province"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label22 
         BackColor       =   &H00FFFFFF&
         Caption         =   "City Name"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   3840
         Width           =   1935
      End
      Begin VB.Label Label21 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Address Line 02"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   3480
         Width           =   1935
      End
      Begin VB.Label Label20 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Addres Line 01"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   3120
         Width           =   1935
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Address Category"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Address- 03"
         Height          =   255
         Left            =   9960
         TabIndex        =   29
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Address - 02"
         Height          =   255
         Left            =   6120
         TabIndex        =   28
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Address - 01"
         Height          =   255
         Left            =   2160
         TabIndex        =   27
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.TextBox txtEmpNo 
      Height          =   285
      Left            =   2280
      TabIndex        =   3
      Top             =   840
      Width           =   1575
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   9900
      Width           =   14880
      _ExtentX        =   26247
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
   Begin VB.Label Label28 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Contact Number - 02"
      Height          =   255
      Left            =   5160
      TabIndex        =   40
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label Label19 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Contact Number - 01"
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label Label13 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   11040
      TabIndex        =   23
      Top             =   3000
      Width           =   3735
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Gender"
      Height          =   255
      Left            =   10080
      TabIndex        =   22
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label lblCat02 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   7200
      TabIndex        =   21
      Top             =   3000
      Width           =   2775
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Category - 02"
      Height          =   255
      Left            =   5160
      TabIndex        =   20
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label lblCat01 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2280
      TabIndex        =   19
      Top             =   3000
      Width           =   2775
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Category - 01"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label lblDepartment 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2280
      TabIndex        =   17
      Top             =   1560
      Width           =   12495
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Department"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label lblFactory 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   8040
      TabIndex        =   15
      Top             =   840
      Width           =   6735
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Factory"
      Height          =   255
      Left            =   6840
      TabIndex        =   14
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2280
      TabIndex        =   13
      Top             =   2640
      Width           =   12495
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Designation"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label lblSubSection 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   2280
      Width           =   12495
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sub Section"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label lblSection 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2280
      TabIndex        =   9
      Top             =   1920
      Width           =   12495
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Section"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label FullName 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   1200
      Width           =   12495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Employee Name"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label lblEPF 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5160
      TabIndex        =   5
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "EPF Number"
      Height          =   255
      Left            =   3960
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Employee Number"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   0
      Picture         =   "frmExtEmpMSTR.frx":2A4A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Employee Master"
      BeginProperty Font 
         Name            =   "Neuropolitical Rg"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   120
      Width           =   8055
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
End
Attribute VB_Name = "frmExtEmpMSTR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cmbProvince_ID, cmbDistrict_ID As Integer
Dim District_Name, Province_Name As String
Private Sub cmdExit_Click()
    Close All
    Unload Me
End Sub

Private Sub dcmbDistrict1_LostFocus()
    cmbDistrict_ID = 1
    If Trim(dcmbDistrict1.Text) <> "" Then
        PR_Fill_MOH
    End If
End Sub

Private Sub dcmbDistrict2_Click(Area As Integer)
    cmbDistrict_ID = 2
    If Trim(dcmbDistrict2.Text) <> "" Then
        PR_Fill_MOH
    End If
End Sub

Private Sub dcmbDistrict3_Click(Area As Integer)
    cmbDistrict_ID = 3
    If Trim(dcmbDistrict3.Text) <> "" Then
        PR_Fill_MOH
    End If
End Sub

Private Sub dcmbProvince1_LostFocus()
    cmbProvince_ID = 1
    If Trim(dcmbProvince1.Text) <> "" Then
        dcmbDistrict1.Text = ""
        PR_Fill_District
        PR_Fill_Police
    End If
End Sub

Private Sub dcmbProvince2_LostFocus()
    cmbProvince_ID = 2
    If Trim(dcmbProvince2.Text) <> "" Then
        dcmbDistrict32Text = ""
        PR_Fill_District
        PR_Fill_Police
    End If
End Sub
Private Sub dcmbProvince3_LostFocus()
    cmbProvince_ID = 3
    If Trim(dcmbProvince3.Text) <> "" Then
        dcmbDistrict3.Text = ""
        PR_Fill_District
        PR_Fill_Police
    End If
End Sub

Private Sub Form_Load()
    PR_HRS_Open_CON
    PR_Fil_Res_Category
    PR_Fill_Province
End Sub

Public Sub PR_Fil_Res_Category()
    Set RS = New ADODB.Recordset
    RS.Open "Select Cat_Description from HRS_HR_MSTR_Category Where Cat_Code='0001' Order by ID", HRS, adOpenKeyset, adLockReadOnly
    dcmbRCat1.ListField = "Cat_Description"
    Set dcmbRCat1.RowSource = RS
    dcmbRCat2.ListField = "Cat_Description"
    Set dcmbRCat2.RowSource = RS
    dcmbRCat3.ListField = "Cat_Description"
    Set dcmbRCat3.RowSource = RS
End Sub

Public Sub PR_Fill_Province()
    Set RS = New ADODB.Recordset
    RS.Open "Select Cat_Description from HRS_HR_MSTR_Category Where Cat_Code='0002' Order by ID", HRS, adOpenKeyset, adLockReadOnly
    dcmbProvince1.ListField = "Cat_Description"
    Set dcmbProvince1.RowSource = RS
    dcmbProvince2.ListField = "Cat_Description"
    Set dcmbProvince2.RowSource = RS
    dcmbProvince3.ListField = "Cat_Description"
    Set dcmbProvince3.RowSource = RS
End Sub

Public Sub PR_Fill_District()
    Set RS = New ADODB.Recordset
    If cmbProvince_ID = 1 Then
        Province_Name = Trim(dcmbProvince1.Text)
    End If
    If cmbProvince_ID = 2 Then
        Province_Name = Trim(dcmbProvince2.Text)
    End If
    If cmbProvince_ID = 3 Then
        Province_Name = Trim(dcmbProvince3.Text)
    End If
    
    RS.Open "Select district_Name from HRSV_HR_MSTR_District Where Province_Name='" & Province_Name & "'  Order by district_Name", HRS, adOpenKeyset, adLockReadOnly
    
    If cmbProvince_ID = 1 Then
        dcmbDistrict1.ListField = "district_Name"
        Set dcmbDistrict1.RowSource = RS
    End If
    If cmbProvince_ID = 2 Then
        dcmbDistrict2.ListField = "district_Name"
        Set dcmbDistrict2.RowSource = RS
    End If
    If cmbProvince_ID = 3 Then
        dcmbDistrict3.ListField = "district_Name"
        Set dcmbDistrict3.RowSource = RS
    End If
End Sub

Public Sub PR_Fill_Police()
    Set RS = New ADODB.Recordset
    If cmbProvince_ID = 1 Then
        Province_Name = Trim(dcmbProvince1.Text)
    End If
    If cmbProvince_ID = 2 Then
        Province_Name = Trim(dcmbProvince2.Text)
    End If
    If cmbProvince_ID = 3 Then
        Province_Name = Trim(dcmbProvince3.Text)
    End If
    
    RS.Open "Select Police_Station from HRSV_HR_Police Where Province_Name='" & Province_Name & "' Order by Police_Station", HRS, adOpenKeyset, adLockReadOnly
    
    If cmbProvince_ID = 1 Then
        dcmbPolice1.ListField = "Police_Station"
        Set dcmbPolice1.RowSource = RS
    End If
    If cmbProvince_ID = 2 Then
        dcmbPolice2.ListField = "Police_Station"
        Set dcmbPolice2.RowSource = RS
    End If
    If cmbProvince_ID = 3 Then
        dcmbPolice3.ListField = "Police_Station"
        Set dcmbPolice3.RowSource = RS
    End If
End Sub

Public Sub PR_Fill_MOH()
    Set RS = New ADODB.Recordset
    If cmbDistrict_ID = 1 Then
        District_Name = Trim(dcmbDistrict1.Text)
    End If
    If cmbDistrict_ID = 2 Then
        District_Name = Trim(dcmbDistrict2.Text)
    End If
    If cmbDistrict_ID = 3 Then
        District_Name = Trim(dcmbDistrict3.Text)
    End If
    
    RS.Open "Select MOH_Office_City from HRSV_HR_MSTR_MOH Where Province_Name='" & Province_Name & "' and District_Name='" & District_Name & "' Order by MOH_Office_City", HRS, adOpenKeyset, adLockReadOnly
    
    If cmbDistrict_ID = 1 Then
        dcmbMOH1.ListField = "MOH_Office_City"
        Set dcmbMOH1.RowSource = RS
    End If
    If cmbProvince_ID = 2 Then
        dcmbMOH2.ListField = "MOH_Office_City"
        Set dcmbMOH2.RowSource = RS
    End If
    If cmbDistrict_ID = 3 Then
        dcmbMOH3.ListField = "MOH_Office_City"
        Set dcmbMOH3.RowSource = RS
    End If
End Sub

