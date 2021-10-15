VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmVMF 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Vehicle Master File"
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8985
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
   ScaleHeight     =   9120
   ScaleWidth      =   8985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbRMethod 
      Height          =   315
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1200
      Width           =   2055
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   6015
      Left            =   120
      TabIndex        =   45
      Top             =   2760
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   10610
      _Version        =   393216
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "&General Options"
      TabPicture(0)   =   "frmVMF.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame5"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Owner Informations"
      TabPicture(1)   =   "frmVMF.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdOISave"
      Tab(1).Control(1)=   "cmdOIEdit"
      Tab(1).Control(2)=   "cmdOIBlacklist"
      Tab(1).Control(3)=   "cmdOIApprove"
      Tab(1).Control(4)=   "txtOName"
      Tab(1).Control(5)=   "txtOAdd1"
      Tab(1).Control(6)=   "txtONIC"
      Tab(1).Control(7)=   "txtOCont1"
      Tab(1).Control(8)=   "txtOCont2"
      Tab(1).Control(9)=   "cmbOBank"
      Tab(1).Control(10)=   "cmbOBranch"
      Tab(1).Control(11)=   "txtOAccNo"
      Tab(1).Control(12)=   "Shape3"
      Tab(1).Control(13)=   "Image7"
      Tab(1).Control(14)=   "Image6"
      Tab(1).Control(15)=   "ImgOI"
      Tab(1).Control(16)=   "Label13"
      Tab(1).Control(17)=   "Label14"
      Tab(1).Control(18)=   "Label15"
      Tab(1).Control(19)=   "Label16"
      Tab(1).Control(20)=   "Label17"
      Tab(1).Control(21)=   "Label18"
      Tab(1).Control(22)=   "Label19"
      Tab(1).Control(23)=   "Label20"
      Tab(1).ControlCount=   24
      TabCaption(2)   =   "&Driver Informations"
      TabPicture(2)   =   "frmVMF.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label25"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label24"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label23"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label22"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label21"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "ImgDI"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Image9"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Image10"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Shape4"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Label30"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Label31"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "txtDCont2"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "txtDCont1"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "txtDNIC"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "txtDAddress"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "txtDName"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "cmdDIBlackList"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "cmdDIApprove"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "cmdDISave"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "cmdDIEdit"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "txtLSNo"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "DTPDLExp"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).ControlCount=   22
      Begin MSComCtl2.DTPicker DTPDLExp 
         Height          =   375
         Left            =   1800
         TabIndex        =   80
         Top             =   2640
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         Format          =   55443457
         CurrentDate     =   43298
      End
      Begin VB.TextBox txtLSNo 
         Height          =   285
         Left            =   1800
         TabIndex        =   79
         Top             =   2280
         Width           =   2055
      End
      Begin VB.CommandButton cmdDIEdit 
         Enabled         =   0   'False
         Height          =   975
         Left            =   6600
         Picture         =   "frmVMF.frx":0054
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   3600
         Width           =   975
      End
      Begin VB.CommandButton cmdDISave 
         Enabled         =   0   'False
         Height          =   975
         Left            =   6600
         Picture         =   "frmVMF.frx":0D24
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   4680
         Width           =   975
      End
      Begin VB.CommandButton cmdDIApprove 
         Enabled         =   0   'False
         Height          =   975
         Left            =   360
         Picture         =   "frmVMF.frx":156F
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   4680
         Width           =   975
      End
      Begin VB.CommandButton cmdDIBlackList 
         Enabled         =   0   'False
         Height          =   975
         Left            =   360
         Picture         =   "frmVMF.frx":1DCE
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   3600
         Width           =   975
      End
      Begin VB.TextBox txtDName 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   23
         Top             =   480
         Width           =   5775
      End
      Begin VB.TextBox txtDAddress 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   24
         Top             =   840
         Width           =   5775
      End
      Begin VB.TextBox txtDNIC 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   25
         Top             =   1200
         Width           =   5775
      End
      Begin VB.TextBox txtDCont1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   26
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox txtDCont2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   27
         Top             =   1920
         Width           =   2055
      End
      Begin VB.CommandButton cmdOISave 
         Enabled         =   0   'False
         Height          =   975
         Left            =   -68400
         Picture         =   "frmVMF.frx":2723
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   4680
         Width           =   975
      End
      Begin VB.CommandButton cmdOIEdit 
         Enabled         =   0   'False
         Height          =   975
         Left            =   -68400
         Picture         =   "frmVMF.frx":2F6E
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   3600
         Width           =   975
      End
      Begin VB.CommandButton cmdOIBlacklist 
         Enabled         =   0   'False
         Height          =   975
         Left            =   -74760
         Picture         =   "frmVMF.frx":3C3E
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   3600
         Width           =   975
      End
      Begin VB.CommandButton cmdOIApprove 
         Enabled         =   0   'False
         Height          =   975
         Left            =   -74760
         Picture         =   "frmVMF.frx":4593
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   4680
         Width           =   975
      End
      Begin VB.TextBox txtOName 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73200
         TabIndex        =   16
         Top             =   480
         Width           =   5655
      End
      Begin VB.TextBox txtOAdd1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73200
         TabIndex        =   17
         Top             =   840
         Width           =   5655
      End
      Begin VB.TextBox txtONIC 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73200
         TabIndex        =   18
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox txtOCont1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73200
         TabIndex        =   19
         Top             =   1560
         Width           =   2295
      End
      Begin VB.TextBox txtOCont2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73200
         TabIndex        =   20
         Top             =   1920
         Width           =   2295
      End
      Begin VB.ComboBox cmbOBank 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -73200
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   2280
         Width           =   5775
      End
      Begin VB.ComboBox cmbOBranch 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -73200
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   2640
         Width           =   5775
      End
      Begin VB.TextBox txtOAccNo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73200
         TabIndex        =   22
         Top             =   3000
         Width           =   2295
      End
      Begin VB.Frame Frame5 
         Height          =   5535
         Left            =   -74880
         TabIndex        =   46
         Top             =   360
         Width           =   7695
         Begin VB.CommandButton cmdGOEdit 
            Enabled         =   0   'False
            Height          =   975
            Left            =   6480
            Picture         =   "frmVMF.frx":4DF2
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   3240
            Width           =   975
         End
         Begin VB.ComboBox cmbVGrade 
            Height          =   315
            ItemData        =   "frmVMF.frx":5AC2
            Left            =   1800
            List            =   "frmVMF.frx":5AC4
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   2400
            Width           =   1815
         End
         Begin VB.CommandButton cmdGOApprove 
            Enabled         =   0   'False
            Height          =   975
            Left            =   240
            Picture         =   "frmVMF.frx":5AC6
            Style           =   1  'Graphical
            TabIndex        =   57
            Top             =   4320
            Width           =   975
         End
         Begin VB.CommandButton cmdGOSave 
            Enabled         =   0   'False
            Height          =   975
            Left            =   6480
            Picture         =   "frmVMF.frx":6325
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   4320
            Width           =   975
         End
         Begin VB.CommandButton cmdGOBlacklist 
            Enabled         =   0   'False
            Height          =   975
            Left            =   240
            Picture         =   "frmVMF.frx":6B70
            Style           =   1  'Graphical
            TabIndex        =   56
            Top             =   3240
            Width           =   975
         End
         Begin VB.ComboBox cmbRoof 
            Height          =   315
            ItemData        =   "frmVMF.frx":74C5
            Left            =   5400
            List            =   "frmVMF.frx":74D2
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   600
            Width           =   2055
         End
         Begin VB.TextBox txtRemarks 
            Height          =   285
            Left            =   1800
            TabIndex        =   13
            Top             =   2760
            Width           =   5655
         End
         Begin VB.ComboBox cmbInscompany 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   960
            Width           =   5655
         End
         Begin VB.ComboBox cmbAC 
            Height          =   315
            ItemData        =   "frmVMF.frx":74F1
            Left            =   1800
            List            =   "frmVMF.frx":7501
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   600
            Width           =   1815
         End
         Begin VB.TextBox txtVSeatCap 
            Height          =   285
            Left            =   1800
            TabIndex        =   5
            Top             =   240
            Width           =   855
         End
         Begin MSComCtl2.DTPicker DTPLExp 
            Height          =   255
            Left            =   1800
            TabIndex        =   10
            Top             =   1680
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   450
            _Version        =   393216
            Format          =   55443457
            CurrentDate     =   43272
         End
         Begin MSComCtl2.DTPicker DTPInsExp 
            Height          =   255
            Left            =   1800
            TabIndex        =   9
            Top             =   1320
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   450
            _Version        =   393216
            Format          =   55443457
            CurrentDate     =   43272
         End
         Begin MSComCtl2.DTPicker DTPAudit 
            Height          =   255
            Left            =   1800
            TabIndex        =   11
            Top             =   2040
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   450
            _Version        =   393216
            Format          =   55443457
            CurrentDate     =   43272
         End
         Begin VB.Shape Shape2 
            Height          =   2295
            Left            =   1800
            Top             =   3120
            Width           =   4215
         End
         Begin VB.Image Image5 
            BorderStyle     =   1  'Fixed Single
            Height          =   2295
            Left            =   120
            Picture         =   "frmVMF.frx":7523
            Stretch         =   -1  'True
            Top             =   3120
            Width           =   1215
         End
         Begin VB.Image Image4 
            BorderStyle     =   1  'Fixed Single
            Height          =   2295
            Left            =   6360
            Picture         =   "frmVMF.frx":8D58
            Stretch         =   -1  'True
            Top             =   3120
            Width           =   1215
         End
         Begin VB.Image ImgGO 
            Height          =   2055
            Left            =   1920
            Picture         =   "frmVMF.frx":A58D
            Stretch         =   -1  'True
            Top             =   3240
            Visible         =   0   'False
            Width           =   3975
         End
         Begin VB.Label Label26 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Roof Type"
            Height          =   255
            Left            =   3720
            TabIndex        =   55
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label12 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Remarks"
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   2760
            Width           =   1575
         End
         Begin VB.Label Label11 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Vehicle Grade"
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   2400
            Width           =   1575
         End
         Begin VB.Label Label10 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Last Audit Date"
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   2040
            Width           =   1575
         End
         Begin VB.Label Label9 
            BackColor       =   &H00E0E0E0&
            Caption         =   "License Expiry Date"
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label Label8 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Insurence Expiry Date"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label Label7 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Insurence Company"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label Label6 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Air Condition Status"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label5 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Seating Capacity"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Label Label31 
         BackColor       =   &H00E0E0E0&
         Caption         =   "License Expiry Date"
         Height          =   255
         Left            =   120
         TabIndex        =   78
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label30 
         BackColor       =   &H00E0E0E0&
         Caption         =   "License Number"
         Height          =   255
         Left            =   120
         TabIndex        =   77
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Shape Shape4 
         Height          =   2535
         Left            =   1560
         Top             =   3240
         Width           =   4815
      End
      Begin VB.Shape Shape3 
         Height          =   2415
         Left            =   -73200
         Top             =   3360
         Width           =   4575
      End
      Begin VB.Image Image10 
         BorderStyle     =   1  'Fixed Single
         Height          =   2295
         Left            =   6480
         Picture         =   "frmVMF.frx":F815
         Stretch         =   -1  'True
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Image Image9 
         BorderStyle     =   1  'Fixed Single
         Height          =   2295
         Left            =   240
         Picture         =   "frmVMF.frx":1104A
         Stretch         =   -1  'True
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Image ImgDI 
         Height          =   2295
         Left            =   1680
         Picture         =   "frmVMF.frx":1287F
         Stretch         =   -1  'True
         Top             =   3360
         Visible         =   0   'False
         Width           =   4575
      End
      Begin VB.Label Label21 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Driver Name"
         Height          =   255
         Left            =   120
         TabIndex        =   74
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label22 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Driverr Address"
         Height          =   255
         Left            =   120
         TabIndex        =   73
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label23 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Driver NIC Number"
         Height          =   255
         Left            =   120
         TabIndex        =   72
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label24 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Driver Contact No 01"
         Height          =   255
         Left            =   120
         TabIndex        =   71
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label25 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Driver Contact No 02"
         Height          =   255
         Left            =   120
         TabIndex        =   70
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Image Image7 
         BorderStyle     =   1  'Fixed Single
         Height          =   2295
         Left            =   -68520
         Picture         =   "frmVMF.frx":17B07
         Stretch         =   -1  'True
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Image Image6 
         BorderStyle     =   1  'Fixed Single
         Height          =   2295
         Left            =   -74880
         Picture         =   "frmVMF.frx":1933C
         Stretch         =   -1  'True
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Image ImgOI 
         Height          =   2175
         Left            =   -73080
         Picture         =   "frmVMF.frx":1AB71
         Stretch         =   -1  'True
         Top             =   3480
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.Label Label13 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Owner Name"
         Height          =   255
         Left            =   -74880
         TabIndex        =   65
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label14 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Owner Address"
         Height          =   255
         Left            =   -74880
         TabIndex        =   64
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label15 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Owner NIC Number"
         Height          =   255
         Left            =   -74880
         TabIndex        =   63
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label16 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Owner Contact No 01"
         Height          =   255
         Left            =   -74880
         TabIndex        =   62
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label17 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Owner Contact No 02"
         Height          =   255
         Left            =   -74880
         TabIndex        =   61
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label18 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Bank Name"
         Height          =   255
         Left            =   -74880
         TabIndex        =   60
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label19 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Branch Name"
         Height          =   255
         Left            =   -74880
         TabIndex        =   59
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label20 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Account Number"
         Height          =   255
         Left            =   -74880
         TabIndex        =   58
         Top             =   3000
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdProcess 
      Height          =   615
      Left            =   8280
      Picture         =   "frmVMF.frx":1FDF9
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   7440
      Width           =   495
   End
   Begin VB.CommandButton cmdEdit 
      Enabled         =   0   'False
      Height          =   975
      Left            =   6720
      Picture         =   "frmVMF.frx":20400
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Height          =   975
      Left            =   4320
      Picture         =   "frmVMF.frx":210D0
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdate 
      Enabled         =   0   'False
      Height          =   975
      Left            =   7800
      Picture         =   "frmVMF.frx":2191B
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   1320
      Width           =   975
   End
   Begin VB.ComboBox cmbVCat 
      Height          =   315
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox txtVno 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2040
      TabIndex        =   0
      Top             =   840
      Width           =   2055
   End
   Begin VB.ComboBox cmbVBrand 
      Height          =   315
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1920
      Width           =   2055
   End
   Begin VB.ComboBox cmbVModel 
      Height          =   315
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2280
      Width           =   2055
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   43
      Top             =   8865
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   10
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Enabled         =   0   'False
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "10/17/2020"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "12:53 AM"
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
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   6015
      Left            =   8160
      TabIndex        =   44
      Top             =   2760
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
         Picture         =   "frmVMF.frx":22605
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Exit from the System"
         Top             =   5400
         Width           =   495
      End
   End
   Begin VB.Label Label32 
      Caption         =   "Registered Method"
      Height          =   255
      Left            =   120
      TabIndex        =   81
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Left            =   4200
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   4695
   End
   Begin VB.Label Label27 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Save"
      Height          =   255
      Left            =   4320
      TabIndex        =   42
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Edit"
      Height          =   255
      Left            =   6720
      TabIndex        =   41
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Update"
      Height          =   255
      Left            =   7800
      TabIndex        =   40
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Vehicle Category"
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Vehicle Number"
      Height          =   255
      Left            =   120
      TabIndex        =   35
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Vehicle Brand Name"
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Vehicle Model"
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   0
      Picture         =   "frmVMF.frx":22E4A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2175
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   2280
      Picture         =   "frmVMF.frx":23F42
      Stretch         =   -1  'True
      Top             =   120
      Width           =   4575
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
End
Attribute VB_Name = "frmVMF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public V_No As String
Public BL As Integer
Public BL_Mode As String
Public F_Load As Integer
Public Bank_Name As String

Private Sub cmbOBank_Click()
    Bank_Name = Trim(cmbOBank.Text)
    PR_Fill_Branch
End Sub

Private Sub cmbVBrand_Click()
    V_Brand_Name = Trim(cmbVBrand.Text)
    PR_Fill_Model
End Sub

Private Sub cmbVCat_Click()
    V_Category = Trim(cmbVCat.Text)
    PR_Fill_Model
End Sub

Private Sub cmdDIApprove_Click()
    BL = 0
    BL_Mode = "D_BL"
    PR_Black_List
    PR_Blaclist_Status
End Sub

Private Sub cmdDIBlackList_Click()
    BL = 1
    BL_Mode = "D_BL"
    PR_Black_List
    PR_Blaclist_Status
End Sub

Private Sub cmdDIEdit_Click()
    PR_DI_Enable
    cmdDIEdit.Enabled = False
    cmdDISave.Enabled = True
    txtDName.SetFocus
    PR_DI_Enable
End Sub

Private Sub cmdDISave_Click()
   Dim D_Name, D_Add1, D_NIC, D_Cont1, D_cont2, D_LNo, D_LS_Exp As String

    If Trim(txtDName.Text) = "" Then
        MsgBox "Driver Name NOT Entered", vbExclamation
        txtDName.SetFocus
        Exit Sub
    End If
    
    If Trim(txtDAddress.Text) = "" Then
        MsgBox "Driver Address NOT Entered", vbExclamation
        txtDAddress.SetFocus
        Exit Sub
    End If
    
    If Trim(txtDNIC.Text) = "" Then
        MsgBox "Driver NIC NOT Entered", vbExclamation
        txtDNIC.SetFocus
        Exit Sub
    End If
    
    If Trim(txtDCont1.Text) = "" Then
        MsgBox "Driver Contact Number 01 NOT Entered", vbExclamation
        txtDCont1.SetFocus
        Exit Sub
    End If
    
    If Trim(txtDCont2.Text) = "" Then
        MsgBox "Driver Contact Number 02 NOT Entered", vbExclamation
        txtDCont2.SetFocus
        Exit Sub
    End If
    
    If Trim(cmbOBank.Text) = "-Select-" Then
        MsgBox "Owner Bank Name NOT Entered", vbExclamation
        cmbOBank.SetFocus
        Exit Sub
    End If
    
    If Trim(txtLSNo.Text) = "" Then
        MsgBox "Driver License Number NOT Entered", vbExclamation
        txtLSNo.SetFocus
        Exit Sub
    End If
    
    D_Name = Replace(Trim(txtDName.Text), "'", "-")
    D_Add1 = Replace(Trim(txtDAddress.Text), "'", "-")
    D_NIC = Replace(Trim(txtDNIC.Text), "'", "-")
    D_Cont1 = Replace(Trim(txtDCont1.Text), "'", "-")
    D_cont2 = Replace(Trim(txtDCont2.Text), "'", "-")
    D_LNo = Replace(Trim(txtLSNo.Text), "'", "-")
    D_LS_Exp = DTPDLExp.Value

    Set UPRS = New ADODB.Recordset
    PR_HRS_Open_CON
    UPRS.Open "Select * from HRSV_TR_MSTR_Vehicles_App Where V_No='" & V_No & "'", HRS, adOpenStatic, adLockOptimistic
    If UPRS.EOF = False Then
        UPRS.Fields("D_Name") = D_Name
        UPRS.Fields("D_Address") = D_Add1
        UPRS.Fields("D_NIC") = D_NIC
        UPRS.Fields("D_Cont1") = D_Cont1
        UPRS.Fields("D_cont2") = D_cont2
        UPRS.Fields("D_LicenseNo") = D_LNo
        UPRS.Fields("D_License_Exp") = D_LS_Exp
        UPRS.Update
    End If
    MsgBox "Vehicle Driver Details Saved Successfully", vbInformation
    cmdDISave.Enabled = False
    cmdDIEdit.Enabled = True
    PR_DI_Dissable
    PR_HRS_Close_CON
End Sub

Private Sub cmdEdit_Click()
    PR_V_Enable
    cmdEdit.Enabled = False
    cmdUpdate.Enabled = True
End Sub

Private Sub cmdExit_Click()
    Close All
    Unload Me
End Sub

Private Sub cmdGOApprove_Click()
    BL = 0
    BL_Mode = "V_BL"
    PR_Black_List
    PR_Blaclist_Status
End Sub

Private Sub cmdGOBlacklist_Click()
    BL = 1
    BL_Mode = "V_BL"
    PR_Black_List
    PR_Blaclist_Status
End Sub

Private Sub cmdGOEdit_Click()
    PR_GO_Enable
    cmdGOEdit.Enabled = False
    cmdGOSave.Enabled = True
    PR_Blaclist_Status
    txtVSeatCap.SetFocus
End Sub

Private Sub cmdGOSave_Click()
    If Trim(txtVSeatCap.Text) = "" Then
        MsgBox "Vehicle Seating Capacity NOT Entered", vbExclamation
        txtVSeatCap.SetFocus
        Exit Sub
    End If
    
    If Trim(cmbAC.Text) = "" Then
        MsgBox "Vehicle AC Type NOT Selected", vbExclamation
        cmbAC.SetFocus
        Exit Sub
    End If

    If Trim(cmbRoof.Text) = "" Then
        MsgBox "Vehicle Roof Type NOT Selected", vbExclamation
        cmbRoof.SetFocus
        Exit Sub
    End If
    
    If Trim(cmbInscompany.Text) = "" Then
        MsgBox "Vehicle Insurence Company NOT Selected", vbExclamation
        cmbInscompany.SetFocus
        Exit Sub
    End If
    
    
    SeatCap = Val(txtVSeatCap.Text)
    AC = Trim(cmbAC.Text)
    Roof = Trim(cmbRoof.Text)
    VGrade = Trim(cmbVGrade.Text)
    IExp = DTPInsExp.Value
    LExp = DTPLExp.Value
    Audit = DTPAudit.Value
    GORemarks = Trim(txtRemarks.Text)
    Ins_Company = Trim(cmbInscompany.Text)
    FN_Find_InsCompany_ID Ins_Company
    Set UPRS = New ADODB.Recordset
    PR_HRS_Open_CON
    UPRS.Open "Select * from HRSV_TR_MSTR_Vehicles Where V_No='" & V_No & "'", HRS, adOpenStatic, adLockOptimistic
    If UPRS.EOF = False Then
        UPRS.Fields("Seat_Cap") = SeatCap
        UPRS.Fields("AC_Type") = AC
        UPRS.Fields("Roof") = Roof
        UPRS.Fields("InsComp_ID") = Ins_Company_ID
        UPRS.Fields("Ins_Exp") = IExp
        UPRS.Fields("Lic_Exp") = LExp
        UPRS.Fields("Aud_Exp") = Audit
        UPRS.Fields("V_Grade") = VGrade
        UPRS.Fields("Remarks") = GORemarks
        UPRS.Update
    End If
    MsgBox "Vehicle General Options Saved Successfully", vbInformation
    cmdGOSave.Enabled = False
    cmdGOEdit.Enabled = True
    PR_GO_Dissable
    PR_HRS_Close_CON
End Sub

Private Sub cmdOIApprove_Click()
    BL = 0
    BL_Mode = "O_BL"
    PR_Black_List
    PR_Blaclist_Status
End Sub

Private Sub cmdOIBlacklist_Click()
    BL = 1
    BL_Mode = "O_BL"
    PR_Black_List
    PR_Blaclist_Status
End Sub

Private Sub cmdOIEdit_Click()
    PR_OI_Enable
    cmdOIEdit.Enabled = False
    cmdOISave.Enabled = True
    txtOName.SetFocus
End Sub

Private Sub cmdOISave_Click()
    Dim O_Name, O_Add1, O_NIC, O_Cont1, O_cont2, O_BankCode, O_BranchCode, O_ACCNo As String

    If Trim(txtOName.Text) = "" Then
        MsgBox "Owner Name NOT Entered", vbExclamation
        txtOName.SetFocus
        Exit Sub
    End If
    
    If Trim(txtOAdd1.Text) = "" Then
        MsgBox "Owner Address NOT Entered", vbExclamation
        txtOAdd1.SetFocus
        Exit Sub
    End If
    
    If Trim(txtONIC.Text) = "" Then
        MsgBox "Owner NIC NOT Entered", vbExclamation
        txtONIC.SetFocus
        Exit Sub
    End If
    
    If Trim(txtOCont1.Text) = "" Then
        MsgBox "Owner Contact Number 01 NOT Entered", vbExclamation
        txtOCont1.SetFocus
        Exit Sub
    End If
    
    If Trim(txtOCont2.Text) = "" Then
        MsgBox "Owner Contact Number 02 NOT Entered", vbExclamation
        txtOCont2.SetFocus
        Exit Sub
    End If
    
    If Trim(cmbOBank.Text) = "-Select-" Then
        MsgBox "Owner Bank Name NOT Entered", vbExclamation
        cmbOBank.SetFocus
        Exit Sub
    Else
        If Trim(cmbOBranch.Text) = "-Select-" Then
            MsgBox "Owner Bank Branch Name NOT Entered", vbExclamation
            cmbOBranch.SetFocus
            Exit Sub
        Else
            Set RS = New ADODB.Recordset
            PR_HRS_Open_CON
            RS.Open "Select * from HRSV_sys_MSTR_Banks Where Bank_Name='" & Bank_Name & "' and Branch_Name='" & Trim(cmbOBranch.Text) & "'", HRS, adOpenStatic, adLockReadOnly
            If RS.EOF = False Then
                O_BranchCode = Trim(RS!Branch_Code)
                O_BankCode = Trim(RS!Bank_Code)
            Else
                MsgBox "Bank or Branch NOT Found", vbExclamation
                Exit Sub
            End If
            PR_HRS_Close_CON
        End If
    End If
    
    O_Name = Replace(Trim(txtOName.Text), "'", "-")
    O_Add1 = Replace(Trim(txtOAdd1.Text), "'", "-")
    O_NIC = Replace(Trim(txtONIC.Text), "'", "-")
    O_Cont1 = Replace(Trim(txtOCont1.Text), "'", "-")
    O_cont2 = Replace(Trim(txtOCont2.Text), "'", "-")
    O_ACCNo = Replace(Trim(txtOAccNo.Text), "'", "-")
    

    Set UPRS = New ADODB.Recordset
    PR_HRS_Open_CON
    UPRS.Open "Select * from HRSV_TR_MSTR_Vehicles Where V_No='" & V_No & "'", HRS, adOpenStatic, adLockOptimistic
    If UPRS.EOF = False Then
        UPRS.Fields("O_Name") = O_Name
        UPRS.Fields("O_Address") = O_Add1
        UPRS.Fields("O_NIC") = O_NIC
        UPRS.Fields("O_Cont1") = O_Cont1
        UPRS.Fields("O_cont2") = O_cont2
        UPRS.Fields("O_Bank_Code") = O_BankCode
        UPRS.Fields("O_Brach_Code") = O_BranchCode
        UPRS.Fields("O_Acc_No") = O_ACCNo
        UPRS.Update
    End If
    MsgBox "Vehicle Owner Details Saved Successfully", vbInformation
    cmdOISave.Enabled = False
    cmdOIEdit.Enabled = True
    PR_OI_Dissable
    PR_HRS_Close_CON
End Sub

Private Sub cmdProcess_Click()
    If Trim(txtVno.Text) = "" Then
        MsgBox "Vehicle Number NOT Found", vbExclamation
        txtVno.SetFocus
    Else
        PR_HRS_Open_CON
        Set UPRS = New ADODB.Recordset
        UPRS.Open "Update HRS_TR_MSTR_Vehicles set Confirmed=1 where V_No='" & V_No & "'", HRS, adOpenStatic, adLockOptimistic
        MsgBox "Vehicle Data Confirmed Successfully", vbInformation
        txtVno.Text = ""
        PR_Header_Edit
        PR_GO_Edit
        PR_OI_Edit
        PR_DI_Edit
        PR_Blaclist_Status
        PR_HRS_Close_CON
    End If
End Sub

Private Sub cmdSave_Click()
On Error GoTo er_EH:
    If Trim(txtVno.Text) = "" Then
        MsgBox "Vehicle Number NOT Entered", vbExclamation
        txtVno.SetFocus
        Exit Sub
    End If
    
    If Trim(cmbVCat.Text) = "-Select-" Then
        MsgBox "Vehicle Category NOT Selected", vbExclamation
        cmbVCat.SetFocus
        Exit Sub
    End If
    
    If Trim(cmbVBrand.Text) = "-Select-" Then
        MsgBox "Vehicle Brand Name NOT Selected", vbExclamation
        cmbVBrand.SetFocus
        Exit Sub
    End If
    
    If Trim(cmbVModel.Text) = "-Select-" Then
        MsgBox "Vehicle Model Name NOT Selected", vbExclamation
        cmbVModel.SetFocus
        Exit Sub
    End If
    
    If Trim(cmbRMethod.Text) = "-Select-" Then
        MsgBox "Vehicle registration Module NOT Selected", vbExclamation
        cmbRMethod.SetFocus
        Exit Sub
    End If
    
    V_No = Trim(txtVno.Text)
    V_Category = Trim(cmbVCat.Text)
    FN_Find_VCat_ID V_Category
    V_Brand_Name = Trim(cmbVBrand.Text)
    V_Model = Trim(cmbVModel.Text)
    FN_Find_Model_ID V_Category, V_Brand_Name, V_Model
    
    PR_HRS_Open_CON
    Set DupRS = New ADODB.Recordset
    DupRS.Open "Select * from HRS_TR_MSTR_Vehicles Where V_No='" & V_No & "'", HRS, adOpenStatic, adLockReadOnly
    If DupRS.EOF = False Then
        MsgBox "Duplicate Entry", vbExclamation
        Exit Sub
    Else
        HRS.Execute "INSERT INTO HRS_TR_MSTR_Vehicles(V_No,Model_ID,Seat_Cap,AC_Type,Roof,InsComp_ID,V_Grade,Remarks,Ins_Exp,Lic_Exp,Aud_Exp,Confirmed,Module_ID,Sub_Module_ID)" _
            + "Values('" & V_No & "'," & V_Model_ID & ",0,'N/A','N/A',0,'Z','N/A','" & Format(Date, "MM/dd/yyyy") & "','" & Format(Date, "MM/dd/yyyy") & "','" & Format(Date, "MM/dd/yyyy") & "',0," & Module_ID & "," & Sub_Module_ID & ")"
        
        cmdSave.Enabled = False
        PR_Header_Enable_False
        cmdGOEdit.Enabled = True
        cmdOIEdit.Enabled = True
        cmdDIEdit.Enabled = True
        PR_HRS_Close_CON
    End If
    Exit Sub
    
er_EH:
    MsgBox Err.Description, vbCritical
    Err.Clear
End Sub

Private Sub cmdUpdate_Click()
    If Trim(cmbVCat.Text) = "-Select-" Then
        MsgBox "Vehicle Category NOT Selected", vbExclamation
        cmbVCat.SetFocus
        Exit Sub
    End If
    
    If Trim(cmbVBrand.Text) = "-Select-" Then
        MsgBox "Vehicle Brand Name NOT Selected", vbExclamation
        cmbVBrand.SetFocus
        Exit Sub
    End If
    
    If Trim(cmbVModel.Text) = "-Select-" Then
        MsgBox "Vehicle Model NOT Selected", vbExclamation
        cmbVModel.SetFocus
        Exit Sub
    End If
    
    If Trim(cmbRMethod.Text) = "" Then
        MsgBox "Vehicle Registration Method NOT Selected", vbExclamation
        cmbRMethod.SetFocus
        Exit Sub
    End If
    
    V_Category = Trim(cmbVCat.Text)
    V_Brand_Name = Trim(cmbVBrand.Text)
    V_Model = Trim(cmbVModel.Text)
    
    FN_Find_VCat_ID V_Category
    FN_Find_Model_ID V_Category, V_Brand_Name, V_Model
   
    Set UPRS = New ADODB.Recordset
    PR_HRS_Open_CON
    UPRS.Open "Select * from HRS_TR_MSTR_Vehicles Where V_No='" & V_No & "'", HRS, adOpenStatic, adLockOptimistic
    If UPRS.EOF = False Then
        UPRS.Fields("Model_ID") = V_Model_ID
        UPRS.Fields("Module_ID") = Module_ID
        UPRS.Update
    End If
    PR_HRS_Close_CON
    cmdUpdate.Enabled = False
    PR_V_Enable
    cmdEdit.Enabled = True
End Sub

Private Sub Command4_Click()

End Sub

Private Sub Form_Load()
    PR_Fill_Combo
    
    'FILL VEHICLE GRADE ---------------------
    cmbVGrade.AddItem "Green"
    cmbVGrade.AddItem "Amber"
    cmbVGrade.AddItem "Red"
    
    'SET UP DEFAULT VALUES ------------------
    DTPInsExp.Value = Date
    DTPLExp.Value = Date
    DTPAudit.Value = Date
    cmbVCat.Text = "-Select-"
    cmbVBrand.Text = "-Select-"
    cmbVModel.Text = "-Select-"
    cmbInscompany.Text = "-Select-"
    cmbVGrade.Text = "Green"
    cmbOBank.Text = "-Select-"
End Sub

Public Sub PR_Fill_Combo()
    PR_HRS_Open_CON
    cmbVCat.Clear
    Set RS = New ADODB.Recordset
    RS.Open "Select Category from HRS_TR_MSTR_Category Order by Category", HRS, adOpenStatic, adLockReadOnly
    Do While RS.EOF = False
        cmbVCat.AddItem Trim(RS!Category)
        RS.MoveNext
    Loop

    Set RS = New ADODB.Recordset
    RS.Open "Select Brand_Name from HRS_TR_MSTR_Brand Order by Brand_Name", HRS, adOpenStatic, adLockReadOnly
    Do While RS.EOF = False
        cmbVBrand.AddItem Trim(RS!Brand_Name)
        RS.MoveNext
    Loop

    cmbInscompany.Clear
    Set RS = New ADODB.Recordset
    RS.Open "Select Ins_Name from HRS_TR_MSTR_Insurance Order by Ins_Name", HRS, adOpenStatic, adLockReadOnly
    Do While RS.EOF = False
        cmbInscompany.AddItem Trim(RS!Ins_Name)
        RS.MoveNext
    Loop
    
    cmbOBank.Clear
    Set RS = New ADODB.Recordset
    RS.Open "Select Bank_Name from HRS_sys_MSTR_Banks Order by Bank_Name", HRS, adOpenStatic, adLockReadOnly
    Do While RS.EOF = False
        cmbOBank.AddItem Trim(RS!Bank_Name)
        RS.MoveNext
    Loop
    
    cmbRMethod.Clear
    Set RS = New ADODB.Recordset
    RS.Open "Select Sub_Module_Name from HRSV_sys_MSTR_Modules where Main_Module_Name='Transport' Order by Sub_Module_Name", HRS, adOpenStatic, adLockReadOnly
    Do While RS.EOF = False
        cmbRMethod.AddItem Trim(RS!Sub_Module_Name)
        RS.MoveNext
    Loop
    
    
    cmbVCat.AddItem "-Select-"
    cmbVBrand.AddItem "-Select-"
    cmbVModel.AddItem "-Select-"
    cmbInscompany.AddItem "-Select-"
    cmbOBank.AddItem "-Select-"
    PR_HRS_Close_CON
End Sub

Public Sub PR_Fill_Model()
    PR_HRS_Open_CON
    cmbVModel.Clear
    Set RS = New ADODB.Recordset
    RS.Open "Select Model_Name From HRSV_TR_MSTR_Vehicles Where Category='" & V_Category & "' and Brand_Name='" & V_Brand_Name & "'", HRS, adOpenStatic, adLockReadOnly
    If RS.EOF = False Then
        Do While RS.EOF = False
            cmbVModel.AddItem Trim(RS!Model_Name)
            RS.MoveNext
        Loop
        cmbVModel.AddItem "-Select-"
    Else
        cmbVModel.AddItem "-Select-"
    End If
    cmbVModel.Text = "-Select-"
    PR_HRS_Close_CON
End Sub

Public Sub PR_Header_Enable_False()
    txtVno.Enabled = False
    cmbVCat.Enabled = False
    cmbVBrand.Enabled = False
    cmbVModel.Enabled = False
    cmdSave.Enabled = False
End Sub

Private Sub txtVno_Change()
    V_No = Replace(Trim(txtVno.Text), "'", "-")
    If Len(V_No) > 3 Then
        PR_Header_Edit
        PR_GO_Edit
        PR_OI_Edit
        PR_DI_Edit
        PR_Blaclist_Status
    End If
End Sub

Public Sub PR_Header_Edit()
    PR_HRS_Open_CON
    Dim GRS As ADODB.Recordset
    Set GRS = New ADODB.Recordset
    GRS.Open "Select * from HRSV_TR_MSTR_Vehicles Where V_No='" & V_No & "'", HRS, adOpenStatic, adLockReadOnly
    If GRS.EOF = False Then
        cmbVCat.Text = Trim(GRS!Category)
        PR_HRS_Open_CON
        cmbVBrand.Text = Trim(GRS!Brand_Name)
        cmbVModel.Text = Trim(GRS!Model_Name)
        PR_GO_Dissable
        PR_V_Dissable
        cmdSave.Enabled = False
        cmdEdit.Enabled = True
        cmdGOEdit.Enabled = True
        cmdOIEdit.Enabled = True
        cmdDIEdit.Enabled = True
    Else
        PR_GO_Dissable
        PR_V_Enable
        cmbVCat.Text = "-Select-"
        cmbVBrand.Text = "-Select-"
        cmbVModel.Text = "-Select-"
        cmdSave.Enabled = True
        cmdEdit.Enabled = False
    End If
End Sub

Public Sub PR_V_Enable()
    cmbVCat.Enabled = True
    cmbVBrand.Enabled = True
    cmbVModel.Enabled = True
End Sub

Public Sub PR_V_Dissable()
    cmbVCat.Enabled = False
    cmbVBrand.Enabled = False
    cmbVModel.Enabled = False
End Sub

Public Sub PR_GO_Dissable()
    txtVSeatCap.Enabled = False
    cmbAC.Enabled = False
    cmbRoof.Enabled = False
    cmbInscompany.Enabled = False
    DTPInsExp.Enabled = False
    DTPLExp.Enabled = False
    DTPAudit.Enabled = False
    cmbVGrade.Enabled = False
    txtRemarks.Enabled = False
    cmdGOApprove.Enabled = False
End Sub

Public Sub PR_GO_Enable()
    txtVSeatCap.Enabled = True
    cmbAC.Enabled = True
    cmbRoof.Enabled = True
    cmbInscompany.Enabled = True
    DTPInsExp.Enabled = True
    DTPLExp.Enabled = True
    DTPAudit.Enabled = True
    cmbVGrade.Enabled = True
    txtRemarks.Enabled = True
End Sub
Public Sub PR_GO_Edit()
    Set RS = New ADODB.Recordset
    PR_HRS_Open_CON
    RS.Open "Select * from HRSV_TR_MSTR_Vehicles Where V_No='" & V_No & "'", HRS, adOpenStatic, adLockReadOnly
    If RS.EOF = False Then
        txtVSeatCap.Text = Val(RS!Seat_Cap)
        cmbAC.Text = Trim(RS!AC_Type)
        If IsNull(RS!Roof) = True Then
            cmbRoof.Text = Trim(RS!Roof)
        End If
        If IsNull(RS!Ins_Name) = False Then
            cmbInscompany.Text = Trim(RS!Ins_Name)
        End If
        DTPInsExp.Value = RS!Ins_Exp
        DTPLExp.Value = RS!Lic_Exp
        DTPAudit.Value = RS!Aud_Exp
        If Trim(RS!V_Grade) <> "Z" Then
            cmbVGrade.Text = Trim(RS!V_Grade)
        End If
        If IsNull(RS!Remarks) = False Then
            txtRemarks.Text = Trim(RS!Remarks)
        End If
    Else
        txtVSeatCap.Text = ""
        cmbAC.Text = "N/A"
        cmbRoof.Text = "N/A"
        cmbInscompany.Text = "-Select-"
        DTPInsExp.Value = Date
        DTPLExp.Value = Date
        DTPAudit.Value = Date
        cmbVGrade.Text = "Green"
        txtRemarks.Text = ""
        cmdGOBlacklist.Enabled = False
        cmdGOApprove.Enabled = False
    End If
    PR_HRS_Close_CON
End Sub

Public Sub PR_Black_List()
    PR_HRS_Open_CON
    Set UPRS = New ADODB.Recordset
    UPRS.Open "update HRS_sys_Vehicles set " & BL_Mode & "=" & BL & " Where V_No='" & V_No & "'", HRS, adOpenStatic, adLockPessimistic
    PR_HRS_Close_CON
End Sub

Public Sub PR_Blaclist_Status()
    Set RS = New ADODB.Recordset
    PR_HRS_Open_CON
    RS.Open "Select * from HRSV_TR_MSTR_Vehicles Where V_No='" & V_No & "'", HRS, adOpenStatic, adLockReadOnly
    If RS.EOF = False Then
        If Trim(RS!V_BL) = True Then
            ImgGO.Visible = True
            cmdGOApprove.Enabled = True
            cmdGOBlacklist.Enabled = False
        Else
            ImgGO.Visible = False
            cmdGOApprove.Enabled = False
            cmdGOBlacklist.Enabled = True
        End If
    
        If Trim(RS!O_BL) = True Then
            ImgOI.Visible = True
            cmdOIApprove.Enabled = True
            cmdOIBlacklist.Enabled = False
        Else
            ImgOI.Visible = False
            cmdOIApprove.Enabled = False
            cmdOIBlacklist.Enabled = True
        End If
        
        If Trim(RS!D_BL) = True Then
            ImgDI.Visible = True
            cmdDIApprove.Enabled = True
            cmdDIBlackList.Enabled = False
        Else
            ImgDI.Visible = False
            cmdDIApprove.Enabled = False
            cmdDIBlackList.Enabled = True
        End If
    Else
        ImgGO.Visible = False
        ImgDI.Visible = False
        ImgOI.Visible = False
    End If
    PR_HRS_Close_CON
End Sub

Public Sub PR_OI_Dissable()
    txtOName.Enabled = False
    txtOAdd1.Enabled = False
    txtONIC.Enabled = False
    txtOCont1.Enabled = False
    txtOCont2.Enabled = False
    cmbOBank.Enabled = False
    cmbOBranch.Enabled = False
    txtOAccNo.Enabled = False
End Sub

Public Sub PR_OI_Enable()
    txtOName.Enabled = True
    txtOAdd1.Enabled = True
    txtONIC.Enabled = True
    txtOCont1.Enabled = True
    txtOCont2.Enabled = True
    cmbOBank.Enabled = True
    cmbOBranch.Enabled = True
    txtOAccNo.Enabled = True
End Sub

Public Sub PR_OI_Edit()
    Dim tmp_Var As String
    Set RS = New ADODB.Recordset
    PR_HRS_Open_CON
    RS.Open "Select * from HRSV_TR_MSTR_Vehicles Where V_No='" & V_No & "'", HRS, adOpenStatic, adLockReadOnly
    If RS.EOF = False Then
        If IsNull(RS!O_Name) = False Then
            txtOName.Text = Trim(RS!O_Name)
        End If
        If IsNull(RS!O_Address) = False Then
            txtOAdd1.Text = Trim(RS!O_Address)
        End If
        
        If IsNull(RS!O_NIC) = False Then
            txtONIC.Text = Trim(RS!O_NIC)
        End If
        If IsNull(RS!O_Cont1) = False Then
            txtOCont1.Text = Trim(RS!O_Cont1)
        End If
        If IsNull(RS!O_cont2) = False Then
            txtOCont2.Text = Trim(RS!O_cont2)
        End If
        If IsNull(RS!O_ACC_No) = False Then
            txtOAccNo.Text = Trim(RS!O_ACC_No)
        End If
        cmbOBranch.Enabled = True
        If IsNull(RS!Branch_Name) = False Then
            tmp_Var = Trim(RS!Branch_Name)
        End If
        If IsNull(RS!Bank_Name) = False Then
            cmbOBank.Text = Trim(RS!Bank_Name)
        End If
        If tmp_Var <> "" Then
            cmbOBranch.Text = tmp_Var
            cmbOBranch.Enabled = False
        End If
    Else
        txtOName.Text = ""
        txtOAdd1.Text = ""
        txtONIC.Text = ""
        txtOCont1.Text = ""
        txtOCont2.Text = ""
        If IsNull(RS!O_ACC_No) = False Then
            txtOAccNo.Text = ""
        End If
        cmbOBranch.Enabled = True
        tmp_Var = "-Select-"
        cmbOBank.Text = "-Select-"
        cmbOBranch.Text = tmp_Var
        cmbOBranch.Enabled = False
    End If
    PR_HRS_Close_CON
End Sub

Public Sub PR_Fill_Branch()
    Set RS = New ADODB.Recordset
    PR_HRS_Open_CON
    RS.Open "Select Branch_Name from HRSV_sys_MSTR_Banks Where Bank_Name='" & Bank_Name & "' Order by Branch_Name", HRS, adOpenStatic, adLockReadOnly
    cmbOBranch.Clear
    Do While RS.EOF = False
        cmbOBranch.AddItem Trim(RS!Branch_Name)
        RS.MoveNext
    Loop
    cmbOBranch.AddItem "-Select-"
    cmbOBranch.Text = "-Select-"
    PR_HRS_Close_CON
End Sub

Public Sub PR_DI_Dissable()
    txtDName.Enabled = False
    txtDNIC.Enabled = False
    txtDAddress.Enabled = False
    txtDCont1.Enabled = False
    txtDCont2.Enabled = False
    txtLSNo.Enabled = False
    DTPDLExp.Enabled = False
End Sub


Public Sub PR_DI_Enable()
    txtDName.Enabled = True
    txtDNIC.Enabled = True
    txtDAddress.Enabled = True
    txtDCont1.Enabled = True
    txtDCont2.Enabled = True
    txtLSNo.Enabled = True
    DTPDLExp.Enabled = True
End Sub

Public Sub PR_DI_Edit()
    Dim tmp_Var As String
    Set RS = New ADODB.Recordset
    PR_HRS_Open_CON
    RS.Open "Select * from HRSV_TR_MSTR_Vehicles Where V_No='" & V_No & "'", HRS, adOpenStatic, adLockReadOnly
    If RS.EOF = False Then
        If IsNull(Trim(RS!D_Name)) = False Then
            txtDName.Text = Trim(RS!D_Name)
        Else
            txtDName.Text = ""
        End If
        If IsNull(Trim(RS!D_Address)) = False Then
            txtDAddress.Text = Trim(RS!D_Address)
        Else
            txtDAddress.Text = ""
        End If
        If IsNull(Trim(RS!D_NIC)) = False Then
            txtDNIC.Text = Trim(RS!D_NIC)
        Else
            txtDNIC.Text = ""
        End If
        If IsNull(Trim(RS!D_Cont1)) = False Then
            txtDCont1.Text = Trim(RS!D_Cont1)
        Else
            txtDCont1.Text = ""
        End If
        
        If IsNull(Trim(RS!D_cont2)) = False Then
            txtDCont2.Text = Trim(RS!D_cont2)
        Else
            txtDCont2.Text = ""
        End If
    Else
        txtDName.Text = ""
        txtDAddress.Text = ""
        txtDNIC.Text = ""
        txtDCont1.Text = ""
        txtDCont2.Text = ""
    End If
    PR_HRS_Close_CON
End Sub
