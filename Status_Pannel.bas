Attribute VB_Name = "Status_Pannel"
Public Function FN_Status_Update(ByRef Status_ID As Integer, Request_ID As Long, R_Module_ID As Integer, R_SubModule_ID As Integer)
    PR_HRS_Open_CON
    Set RS = New ADODB.Recordset
    RS.Open "Select * from HRS_sys_Request_Log Where Request_ID=" & Request_ID & "  and Module_ID=" & R_Module_ID & " AND Sub_Module_ID=" & R_SubModule_ID, HRS, adOpenKeyset, adLockReadOnly
    HRS.Execute "INSERT INTO HRS_sys_Request_Log(Module_ID,Sub_Module_ID,Request_ID,Status_ID,U_ID,Trans_Date_Time,Mechine_ID,Login_ID)" _
            + " Values(" & R_Module_ID & "," & R_SubModule_ID & "," & Request_ID & "," & Status_ID & ",'" & UserID & "','" & Format(Date, "MM/dd/yyyy") + " " + Format(Time, "hh:mm:ss") & "','" & VBA.Environ("COMPUTERNAME") & "','" & UCase(Environ("USERNAME")) & "')"
    Set UPRS = New ADODB.Recordset
    UPRS.Open "Update HRS_TR_Request set Status_ID=" & Status_ID & " Where VR_NO=" & Request_ID & "", HRS, adOpenStatic, adLockOptimistic
    PR_HRS_Close_CON
End Function
