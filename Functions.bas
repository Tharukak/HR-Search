Attribute VB_Name = "Functions"
Public Function FN_DivisionID(ByRef Division As String, ByRef Com_Code As String) As String
    Set IDRS = New ADODB.Recordset
    PR_HRS_Open_CON
    IDRS.Open "Select * from HRS_sys_Division where D_name='" & Division_Name & "' and Com_Code='" & Com_Code & "'", HRS, adOpenStatic, adLockReadOnly
    If IDRS.EOF = False Then
        Division_Code = Trim(IDRS!D_Code)
    Else
        MsgBox "Division Name NOT found", vbExclamation
        Exit Function
    End If
    PR_HRS_Close_CON
End Function

Public Function FN_Find_VCat_ID(ByRef V_Category As String) As Integer
    Set IDRS = New ADODB.Recordset
    PR_HRS_Open_CON
    Rate = 0
    IDRS.Open "Select * from HRS_TR_MSTR_Category where Category='" & V_Category & "'", HRS, adOpenStatic, adLockReadOnly
    If IDRS.EOF = False Then
        V_Cat_ID = Val(IDRS!Cat_ID)
        Rate = Val(IDRS!AC_Rate)
    Else
        MsgBox "Vehicle Category NOT found", vbExclamation
        Exit Function
    End If
    PR_HRS_Close_CON
End Function

Public Function FN_Find_Model_ID(ByRef Category As String, ByRef Brand As String, ByRef Model As String) As Integer
    Set IDRS = New ADODB.Recordset
    PR_HRS_Open_CON
    IDRS.Open "Select * from HRSV_TR_MSTR_Vehicles where Category='" & Category & "' and Brand_Name='" & Brand & "' and Model_Name='" & Model & "'", HRS, adOpenStatic, adLockReadOnly
    If IDRS.EOF = False Then
        V_Model_ID = Val(IDRS!Model_ID)
    Else
        MsgBox "Vehicle Model NOT found", vbExclamation
        Exit Function
    End If
    PR_HRS_Close_CON
End Function

Public Function FN_Find_InsCompany_ID(ByRef Ins_Comp As String) As Integer
    Set IDRS = New ADODB.Recordset
    PR_HRS_Open_CON
    IDRS.Open "Select * from HRS_TR_MSTR_Insurance where Ins_Name='" & Ins_Comp & "'", HRS, adOpenStatic, adLockReadOnly
    If IDRS.EOF = False Then
        Ins_Company_ID = Val(IDRS!Ins_ID)
    Else
        MsgBox "Insurance Company NOT found", vbExclamation
        Exit Function
    End If
    PR_HRS_Close_CON
End Function

Public Function FN_Find_Distance(ByRef From_Loc As String, ByRef To_Loc As String) As Integer
    Set IDRS = New ADODB.Recordset
    PR_HRS_Open_CON
    If From_Loc <> "" And To_Loc <> "" Then
        IDRS.Open "Select * from HRS_TR_MSTR_Distance where From_Loc='" & From_Loc & "' and To_City='" & To_Loc & "'", HRS, adOpenStatic, adLockReadOnly
        If IDRS.EOF = False Then
            Distance = Val(IDRS!Distance)
            Loc_City_ID = Val(IDRS!Dist_ID)
        Else
            Loc_City_ID = 0
            MsgBox "Locations NOT Found", vbExclamation
            Exit Function
        End If
    Else
        From_Loc = ""
        To_Loc = ""
    End If
    PR_HRS_Close_CON
End Function

Public Function FN_Find_Reason_Cat_ID(ByRef Reason_Category As String, ComCode As String) As Integer
    Set IDRS = New ADODB.Recordset
    PR_HRS_Open_CON
    IDRS.Open "Select * from HRS_TR_MSTR_Reason_Cat where Module_ID=1 and Sub_Module_ID=1 and Reason_Category='" & Reason_Category & "' and Com_Code='" & ComCode & "'", HRS, adOpenStatic, adLockReadOnly
    If IDRS.EOF = False Then
        R_Method_ID = Val(IDRS!Reason_Cat_ID)
    Else
        MsgBox "Requested Method NOT Found", vbExclamation
        Exit Function
    End If
    PR_HRS_Close_CON
End Function

Public Function FN_Find_Reason_ID(ByRef Reason As String, ComCode As String) As Integer
    Set IDRS = New ADODB.Recordset
    PR_HRS_Open_CON
    IDRS.Open "Select * from HRS_TR_MSTR_Reason where Module_ID=1 and Sub_Module_ID=1 and Reason_Details='" & Reason & "' and Com_Code='" & ComCode & "'", HRS, adOpenStatic, adLockReadOnly
    If IDRS.EOF = False Then
        Reason_ID = Val(IDRS!Reason_ID)
    Else
        MsgBox "Reason NOT Found", vbExclamation
        Exit Function
    End If
    PR_HRS_Close_CON
End Function

Public Function FN_Find_Province_ID(ByRef Province As String) As Integer
    Set IDRS = New ADODB.Recordset
    PR_HRS_Open_CON
    IDRS.Open "Select * from HRS_HR_MSTR_Category where Cat_Code='0002' and Cat_Description='" & Province & "'", HRS, adOpenStatic, adLockReadOnly
    If IDRS.EOF = False Then
        Province_ID = Val(IDRS!ID)
    Else
        MsgBox "Province NOT Found", vbExclamation
        Exit Function
    End If
    PR_HRS_Close_CON
End Function
