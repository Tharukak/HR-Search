VERSION 5.00
Begin VB.Form HRS_Ex 
   Caption         =   "Form1"
   ClientHeight    =   825
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3675
   Icon            =   "HRS_Extream.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   825
   ScaleWidth      =   3675
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "HRS_Ex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public con As ADODB.Connection
Public MaxVersion As Long
Public M_Major_Version As Long
Public M_Minor_Version As Long
Public M_Revision_Version As Long
Public ApVersion As Long
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
    con.Open "Provider=SQLOLEDB.1;Password=SinX@123;Persist Security Info=True;User ID=HRS_Admin;Initial Catalog=HRS_Extreme;Data Source=10.227.241.27"
    
    Dim recrs As ADODB.Recordset
    Set recrs = New ADODB.Recordset
    
    MechineID = VBA.Environ("COMPUTERNAME")

    recrs.Open "Select isnull(max(Major_Version+Minor_Version+Revision_Version),0) as Version from HRS_sys_Log Where MechineName='" & MechineID & "'", con, adOpenStatic, adLockReadOnly
    If recrs.EOF = False Then
        If IsNull(recrs!Version) = False Then
            ApVersion = recrs!Version
        Else
            ApVersion = 0
        End If
    Else
        ApVersion = 0
    End If
    
    Set recrs = New ADODB.Recordset
    recrs.Open "Select  Major_Version,Minor_Version,Revision_Version,max(Major_Version+Minor_Version+Revision_Version) as Version " _
                + "From HRS_sys_Log Group by  Major_Version,Minor_Version,Revision_Version Having max(Major_Version+Minor_Version+Revision_Version)=(Select max(Major_Version+Minor_Version+Revision_Version) from HRS_sys_Log)", con, adOpenStatic, adLockReadOnly
    If recrs.EOF = False Then
        If IsNull(recrs!Version) = False Then
            MaxVersion = recrs!Version
            M_Major_Version = Val(recrs!Major_Version)
            M_Minor_Version = Val(recrs!Minor_Version)
            M_Revision_Version = Val(recrs!Revision_Version)
        Else
            MaxVersion = 0
        End If
    Else
        MaxVersion = 0
    End If
    Dim ScOMMAND As String
    If ApVersion < MaxVersion Then
        If MsgBox("You are using out dated application, Do you want to update your HR-Search-Extream", vbQuestion + vbYesNo, "Automatic Reporter Update") = vbYes Then
            ScOMMAND = "md c:\HR_Search_EX"
            Shell ("cmd.exe /c" & ScOMMAND)
            DoEvents
            
            ScOMMAND = "md c:\HR_Search_EX\Reports"
            Shell ("cmd.exe /c" & ScOMMAND)
            DoEvents
        
            ScOMMAND = "Copy \\10.151.152.20\sw$\EAG\Inter_Dev\HRS\HR_Search_EX\Reports\*.rpt C:\HR_Search_EX\Reports"
            Shell ("cmd.exe /c" & ScOMMAND)
            DoEvents
            
            Sleep 10000
            
            ScOMMAND = "Copy \\10.151.152.20\sw$\EAG\Inter_Dev\HRS\HR_Search_EX\*.exe C:\HR_Search_EX"
            Shell ("cmd.exe /c" & ScOMMAND)
            DoEvents
            
            GetXpOsArchitecture
            If OSBit = "64" Then
                Sleep 2000
                ScOMMAND = "Copy \\10.151.152.20\sw$\EAG\Inter_Dev\HRS\HR_Search_EX\Sys_Files*.ocx C:\Windows\system"
                Shell ("cmd.exe /c" & ScOMMAND)
                DoEvents
            Else
                Sleep 2000
                ScOMMAND = "Copy \\10.151.152.20\sw$\EAG\Inter_Dev\HRS\HR_Search_EX\Sys_Files\*.ocx C:\Windows\SysWOW64"
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
    Dim sysDateTime As String
    sysDateTime = Format(Date, "MM/dd/yyyy") & " " & Format(Time, "hh:mm:ss")
    con.Execute "INSERT INTO HRS_sys_Log(TRANS_DATE_TIME,MECHINENAME,Major_Version,Minor_Version,Revision_Version) VALUES('" & sysDateTime & "','" & VBA.Environ("COMPUTERNAME") & "'," & M_Major_Version & "," & M_Minor_Version & "," & M_Revision_Version & ")"
End Sub

Public Sub PR_Shell()
    Shell "C:\\HR_Search_EX\HRS_APP.exe", vbNormalFocus
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

