VERSION 5.00
Begin VB.Form ProMIS_SX 
   Caption         =   "Form1"
   ClientHeight    =   1320
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4515
   Icon            =   "ProMIS_SX.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1320
   ScaleWidth      =   4515
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "ProMIS_SX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public con As ADODB.Connection
Public MaxVersion As String
Public ApVersion As String
Public MechineID As String
Public OSBit As String
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Sub Form_Load()
On Error GoTo er_EH:
    PR_Update
    Sleep 100

    Close All
    Unload Me
    Exit Sub
er_EH:
    MsgBox Err.Description, vbExclamation
    Close All
    Unload Me
End Sub

Public Sub PR_Update()
    Set con = New ADODB.Connection
    con.Open "Provider=SQLOLEDB.1;Password=sinx@123;Persist Security Info=True;User ID=App_User;Initial Catalog=ProMIS_SX;Data Source=bli-srv-blim3"
    
    Dim recrs As ADODB.Recordset
    Set recrs = New ADODB.Recordset
    
    MechineID = VBA.Environ("COMPUTERNAME")

    recrs.Open "Select max(Version) as Version from ProMIS_SX_sys_Log Where MechineName='" & MechineID & "'", con, adOpenStatic, adLockReadOnly
    If IsNull(recrs!Version) = False Then
        ApVersion = recrs!Version
    Else
        ApVersion = "0"
    End If
    
    Set recrs = New ADODB.Recordset
    recrs.Open "Select max(Version) as Version from ProMIS_SX_sys_Log", con, adOpenStatic, adLockReadOnly
    If IsNull(recrs!Version) = False Then
        MaxVersion = recrs!Version
    Else
        MaxVersion = "0"
    End If
    Dim ScOMMAND As String
    If ApVersion < MaxVersion Then
        If MsgBox("You are using out dated application, Do you want to update your ProMIS-SX", vbQuestion + vbYesNo, "Automatic Reporter Update") = vbYes Then
            ScOMMAND = "md c:\ProMIS_SX"
            Shell ("cmd.exe /c" & ScOMMAND)
            DoEvents
            
            ScOMMAND = "md c:\ProMIS_SX\ProMIS_SX"
            Shell ("cmd.exe /c" & ScOMMAND)
            DoEvents
            
            ScOMMAND = "md c:\ProMIS_SX\ProMIS_SX\ProMIS_SX_sys_Reports"
            Shell ("cmd.exe /c" & ScOMMAND)
            DoEvents
        
            ScOMMAND = "Copy \\10.227.60.11\Applications$\Reporter_Updates\ProMIS_App_Store\ProMIS_Store\ProMIS_SX_sys_Reports\*.rpt C:\ProMIS_SX\ProMIS_SX\ProMIS_SX_sys_Reports"
            Shell ("cmd.exe /c" & ScOMMAND)
            DoEvents
            
            ScOMMAND = "Copy \\10.227.60.11\Applications$\Reporter_Updates\ProMIS_App_Store\ProMIS_Store\ProMIS_SX_sys_Reports\*.rpt C:\ProMIS_SX\ProMIS_SX\ProMIS_SX_sys_Reports"
            Shell ("cmd.exe /c" & ScOMMAND)
            DoEvents
            
            Sleep 10000
            
            ScOMMAND = "Copy \\10.227.60.11\Applications$\Reporter_Updates\ProMIS_App_Store\ProMIS_Store\*.exe C:\ProMIS_SX\ProMIS_SX"
            Shell ("cmd.exe /c" & ScOMMAND)
            DoEvents
            
            GetXpOsArchitecture
            If OSBit = "64" Then
                Sleep 2000
                ScOMMAND = "Copy \\10.227.60.11\Applications$\Reporter_Updates\ProMIS_App_Store\ProMIS_Store\ProMIS_sys_Files\*.ocx C:\Windows\system"
                Shell ("cmd.exe /c" & ScOMMAND)
                DoEvents
            Else
                Sleep 2000
                ScOMMAND = "Copy \\10.227.60.11\Applications$\Reporter_Updates\ProMIS_App_Store\ProMIS_Store\ProMIS_sys_Files\*.ocx C:\Windows\SysWOW64"
                Shell ("cmd.exe /c" & ScOMMAND)
                DoEvents
            End If
            LOG_REGISTER
            
            PR_Shell
            DoEvents
        Else
            Close All
            Exit Sub
            Unload Me
        End If
    Else
        PR_Shell
    End If
End Sub

Public Sub LOG_REGISTER()
    If con.State = adStateClosed Then con.Open
    con.Execute "INSERT INTO ProMIS_SX_sys_Log(TRANS_DATE_TIME,MECHINENAME,Version) VALUES('" & Date + Time & "','" & VBA.Environ("COMPUTERNAME") & "','" & MaxVersion & "')"
End Sub

Public Sub PR_Shell()
    
    Shell "C:\ProMIS_SX\ProMIS_SX\ProMIS_SX_APP.exe", vbNormalFocus
End Sub

Private Function GetXpOsArchitecture() As String
    Dim ComputerSystemSet As Object
    Dim Computer As Object
    Dim SystemType As String

    Set ComputerSystemSet = GetObject("Winmgmts:"). _
        ExecQuery("SELECT * FROM Win32_ComputerSystem")
    For Each Computer In ComputerSystemSet
        SystemType = UCase$(Left$(Trim$(Computer.SystemType), 3))
    Next

    GetXpOsArchitecture = IIf(SystemType = "X86", "32", "64")
    OSBit = GetXpOsArchitecture
End Function

