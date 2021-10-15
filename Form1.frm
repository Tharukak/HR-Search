VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   6015
      Left            =   1560
      TabIndex        =   0
      Top             =   1560
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   10610
      _Version        =   393216
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "&General Options"
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Owner Informations"
      TabPicture(1)   =   "Form1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&Driver Informations"
      TabPicture(2)   =   "Form1.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame4"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame4 
         Height          =   5535
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   7695
         Begin VB.CommandButton Command5 
            Enabled         =   0   'False
            Height          =   975
            Left            =   6480
            Picture         =   "Form1.frx":0054
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   4320
            Width           =   975
         End
         Begin VB.CommandButton Command4 
            Enabled         =   0   'False
            Height          =   975
            Left            =   4320
            Picture         =   "Form1.frx":089F
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   4320
            Width           =   855
         End
         Begin VB.TextBox Text11 
            Height          =   285
            Left            =   1800
            TabIndex        =   29
            Top             =   1680
            Width           =   2055
         End
         Begin VB.TextBox Text10 
            Height          =   285
            Left            =   1800
            TabIndex        =   28
            Top             =   1320
            Width           =   2055
         End
         Begin VB.TextBox Text9 
            Height          =   285
            Left            =   1800
            TabIndex        =   27
            Top             =   960
            Width           =   5775
         End
         Begin VB.TextBox Text8 
            Height          =   285
            Left            =   1800
            TabIndex        =   26
            Top             =   600
            Width           =   5775
         End
         Begin VB.TextBox Text7 
            Height          =   285
            Left            =   1800
            TabIndex        =   25
            Top             =   240
            Width           =   5775
         End
         Begin VB.CommandButton Command6 
            Enabled         =   0   'False
            Height          =   975
            Left            =   5160
            Picture         =   "Form1.frx":11F4
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   4320
            Width           =   975
         End
         Begin VB.Image Image6 
            BorderStyle     =   1  'Fixed Single
            Height          =   1215
            Left            =   4200
            Picture         =   "Form1.frx":1A53
            Stretch         =   -1  'True
            Top             =   4200
            Width           =   3375
         End
         Begin VB.Label Label25 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Driver Contact No 02"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label Label24 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Driver Contact No 01"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label Label23 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Driver NIC Number"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label Label22 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Driverr Address"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label21 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Driver Name"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   1575
         End
         Begin VB.Image ImgDI 
            Height          =   2295
            Left            =   120
            Picture         =   "Form1.frx":3288
            Stretch         =   -1  'True
            Top             =   3120
            Visible         =   0   'False
            Width           =   3975
         End
      End
      Begin VB.Frame Frame3 
         Height          =   5535
         Left            =   -74880
         TabIndex        =   3
         Top             =   360
         Width           =   7695
         Begin VB.CommandButton Command3 
            Enabled         =   0   'False
            Height          =   975
            Left            =   6480
            Picture         =   "Form1.frx":8510
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   4320
            Width           =   975
         End
         Begin VB.CommandButton Command2 
            Enabled         =   0   'False
            Height          =   975
            Left            =   4320
            Picture         =   "Form1.frx":8D5B
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   4320
            Width           =   855
         End
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   1800
            TabIndex        =   12
            Top             =   2760
            Width           =   2295
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   1800
            TabIndex        =   11
            Top             =   2400
            Width           =   5775
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   1800
            TabIndex        =   10
            Top             =   2040
            Width           =   5775
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   1800
            TabIndex        =   9
            Top             =   1680
            Width           =   2295
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   1800
            TabIndex        =   8
            Top             =   1320
            Width           =   2295
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   1800
            TabIndex        =   7
            Top             =   960
            Width           =   2295
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   1800
            TabIndex        =   6
            Top             =   600
            Width           =   5655
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   1800
            TabIndex        =   5
            Top             =   240
            Width           =   5655
         End
         Begin VB.CommandButton Command1 
            Enabled         =   0   'False
            Height          =   975
            Left            =   5160
            Picture         =   "Form1.frx":96B0
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   4320
            Width           =   975
         End
         Begin VB.Image Image5 
            BorderStyle     =   1  'Fixed Single
            Height          =   1215
            Left            =   4200
            Picture         =   "Form1.frx":9F0F
            Stretch         =   -1  'True
            Top             =   4200
            Width           =   3375
         End
         Begin VB.Label Label20 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Branch Name"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   2760
            Width           =   1575
         End
         Begin VB.Label Label19 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Branch Name"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   2400
            Width           =   1575
         End
         Begin VB.Label Label18 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Bank Name"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   2040
            Width           =   1575
         End
         Begin VB.Label Label17 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Owner Contact No 02"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label Label16 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Owner Contact No 01"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label Label15 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Owner NIC Number"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label Label14 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Owner Address"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label13 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Owner Name"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   1575
         End
         Begin VB.Image ImgOI 
            Height          =   2295
            Left            =   120
            Picture         =   "Form1.frx":B744
            Stretch         =   -1  'True
            Top             =   3120
            Visible         =   0   'False
            Width           =   3975
         End
      End
      Begin VB.Frame Frame2 
         Height          =   5535
         Left            =   -74880
         TabIndex        =   1
         Top             =   360
         Width           =   7695
         Begin VB.ComboBox cmbVGrade 
            Height          =   315
            ItemData        =   "Form1.frx":109CC
            Left            =   -840
            List            =   "Form1.frx":109DF
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   -720
            Width           =   1815
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
