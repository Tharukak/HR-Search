VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmRootAllocation 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Root Allocation"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8955
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
   ScaleHeight     =   8910
   ScaleWidth      =   8955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   8415
      Left            =   8160
      TabIndex        =   2
      Top             =   120
      Width           =   735
      Begin VB.CommandButton cmdExit 
         Height          =   495
         Left            =   120
         Picture         =   "frmRootAllocation.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Exit from the System"
         Top             =   7800
         Width           =   495
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   8655
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   10
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
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSFGrid 
      Height          =   8415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   14843
      _Version        =   393216
   End
End
Attribute VB_Name = "frmRootAllocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    Close All
    Unload Me
End Sub

