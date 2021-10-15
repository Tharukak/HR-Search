Attribute VB_Name = "Mod_SMS"
Public Sub FN_SMS(ByRef Phone_No As String, Message As String)
    PR_SMS_Open_CON
    SMS.Execute "INSERT INTO SMS_Trans_Log(App_ID,Phone_No,Message,Create_Date_Time) Values(2,'" & Phone_No & "','" & Message & "','" & Date + Time & "')"
    PR_SMS_Close_CON
End Sub
