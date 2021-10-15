VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMealCount 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Online Meal Count Dashboard"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13350
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
   ScaleHeight     =   4905
   ScaleWidth      =   13350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   4650
      Width           =   13350
      _ExtentX        =   23548
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   22931
            MinWidth        =   22931
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSFBF 
      Height          =   2175
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   3836
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer2 
      Interval        =   5000
      Left            =   720
      Top             =   5640
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   120
      Top             =   5640
   End
   Begin MSFlexGridLib.MSFlexGrid MSFL 
      Height          =   2175
      Left            =   3720
      TabIndex        =   8
      Top             =   2400
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   3836
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSFES 
      Height          =   2175
      Left            =   7320
      TabIndex        =   9
      Top             =   2400
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   3836
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblLastupdate 
      Alignment       =   2  'Center
      BackColor       =   &H00400040&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   615
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   5295
   End
   Begin VB.Image Image1 
      Height          =   4455
      Left            =   10800
      Picture         =   "frmMealCount.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lblDateTime 
      Alignment       =   2  'Center
      BackColor       =   &H00400040&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   615
      Left            =   5520
      TabIndex        =   6
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label lblSnack 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   735
      Left            =   7320
      TabIndex        =   5
      Top             =   1560
      Width           =   3375
   End
   Begin VB.Label lblLunch 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   3720
      TabIndex        =   4
      Top             =   1560
      Width           =   3375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Evening Snack Count"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   495
      Left            =   7320
      TabIndex        =   3
      Top             =   960
      Width           =   3375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lunch Count"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   960
      Width           =   3375
   End
   Begin VB.Label lblBF 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   3375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Breakfast Count"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   3375
   End
End
Attribute VB_Name = "frmMealCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'    ADMS_Connect
End Sub

Private Sub timer1_Timer()
    lblDateTime.Caption = Format(Date, "dd-MMM-yyyy") + "    " + Format(Time, "HH:mm:ss")
    lblDateTime.Refresh
End Sub
Private Sub Timer2_Timer()
'On Error GoTo er_EH:
    StatusBar1.Panels(1).Text = "Synchronizing ..........."
    Dim BF_T_Count, LU_T_Count, ES_T_Count As Long
    PR_HRS_Open_CON
    Set RS = New ADODB.Recordset
    RS.Open "Select * from HRS_HR_Meal_Dashboard Where Sys_Date='" & Date & "' and Cluster_Code='" & U_Com_Code & "'", HRS, adOpenStatic, adLockReadOnly
    MSFBF.Cols = 2
    MSFL.Cols = 2
    MSFES.Cols = 2
    Dim I As Integer
    I = 1
    BF_T_Count = 0
    LU_T_Count = 0
    ES_T_Count = 0
    MSFBF.Rows = 1
    MSFL.Rows = 1
    MSFES.Rows = 1
    MSFBF.Clear
    Do While RS.EOF = False
        MSFBF.Rows = MSFBF.Rows + 1
        MSFL.Rows = MSFL.Rows + 1
        MSFES.Rows = MSFES.Rows + 1
        MSFBF.TextMatrix(I, 0) = Trim(RS!Clock_Location)
        MSFL.TextMatrix(I, 0) = Trim(RS!Clock_Location)
        MSFES.TextMatrix(I, 0) = Trim(RS!Clock_Location)
        If IsNull(RS!BF_Count) = True Then
            MSFBF.TextMatrix(I, 1) = 0
        Else
            MSFBF.TextMatrix(I, 1) = Val(RS!BF_Count)
            BF_T_Count = BF_T_Count + Val(RS!BF_Count)
        End If
        
        If IsNull(RS!LU_Count) = True Then
            MSFL.TextMatrix(I, 1) = 0
        Else
            MSFL.TextMatrix(I, 1) = Val(RS!LU_Count)
            LU_T_Count = LU_T_Count + Val(RS!LU_Count)
        End If
        If IsNull(RS!ES_Count) = True Then
            MSFES.TextMatrix(I, 1) = 0
        Else
            MSFES.TextMatrix(I, 1) = Val(RS!ES_Count)
            ES_T_Count = ES_T_Count + Val(RS!ES_Count)
        End If
        
        'Tot_Count = Tot_Count + Val(RS!BF_Count)
        lblLastupdate.Caption = Format(RS!Last_Update, "dd-MMM-yyyy") + "    " + Format(RS!Last_Update, "HH:mm:ss")
        I = I + 1
        RS.MoveNext
    Loop
    
    lblBF.Caption = BF_T_Count
    lblLunch.Caption = LU_T_Count
    lblSnack.Caption = ES_T_Count
    
    
    MSFBF.TextMatrix(0, 0) = "Plant Name"
    MSFBF.TextMatrix(0, 1) = "Breakfast Count"
    
    MSFL.TextMatrix(0, 0) = "Plant Name"
    MSFL.TextMatrix(0, 1) = "Lunch Count"
    
    MSFES.TextMatrix(0, 0) = "Plant Name"
    MSFES.TextMatrix(0, 1) = "Eve.Snack Count"
    
    MSFBF.ColWidth(1) = 1700
    MSFL.ColWidth(1) = 1700
    MSFES.ColWidth(1) = 1700

    lblLunch.Refresh
    lblSnack.Refresh
    lblBF.Refresh
    StatusBar1.Panels(1).Text = ""
    PR_HRS_Close_CON
    Exit Sub
    
'er_EH:
'    MsgBox Err.Description, vbCritical
'    Err.Clear
End Sub
