Attribute VB_Name = "Security"
Public Function MenuAccess(ByRef Operation_ID As Integer) As Integer
On Error GoTo ER_EH:
    Dim URS As ADODB.Recordset
    Set URS = New ADODB.Recordset
    PR_HRS_Open_CON
    URS.Open "Select * from HRS_sys_Rights Where User_ID='" & UserID & "' and Operation_ID=" & Operation_ID, HRS, adOpenStatic, adLockReadOnly
    If URS.EOF = True Then
        MsgBox "User Profile NOT Created, Contact your HRS-Extream Administrator", vbCritical, "HRS-Extream"
        RightsMode = 0
    Else
        If URS!Allow = True Then
            RightsMode = 1
        Else
            MsgBox "Access Denied, Contact your HRS-Extream Administrator", vbCritical, "HRS-Extream"
            RightsMode = 0
        End If
    End If
    PR_HRS_Close_CON
    Exit Function
ER_EH:
    MsgBox Err.Description, vbCritical
    Err.Clear
End Function
