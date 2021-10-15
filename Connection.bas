Attribute VB_Name = "Connection"
Public Sub db_Connect()
    On Error GoTo er_EH:
    Set HRS = New ADODB.Connection
    HRS.ConnectionString = "Provider=SQLOLEDB.1;Password=SinX@123;Persist Security Info=True;User ID=HRS_Admin;Initial Catalog=HRS_Extreme;Data Source=10.227.241.27"
    HRS.ConnectionTimeout = 0
    
    Exit Sub
er_EH:
    MsgBox Err.Description, vbCritical
    Err.Clear
End Sub

Public Sub PR_HRS_Close_CON()
    db_Connect
    If HRS.State = adStateOpen Then HRS.Close
End Sub

Public Sub PR_HRS_Open_CON()
    db_Connect
    If HRS.State = adStateclose Then HRS.Open
End Sub

Public Sub PR_REPORT_PATH()
    Report_Path = "C:\HR_Search_EX\Reports\"
    DSN_SETTINGS = "dsn=HRS_EX;uid=HRS_Reporter;pwd=welcome@123"
End Sub

Public Sub SMS_Connect()
On Error GoTo er_EH:
    Set SMS = New ADODB.Connection
    SMS.ConnectionString = "Provider=SQLOLEDB.1;Password=welcome@123;Persist Security Info=True;User ID=SMS_User;Initial Catalog=SMS_Gateway;Data Source=10.227.241.27"
    SMS.ConnectionTimeout = 0
    Exit Sub
er_EH:
    MsgBox Err.Description, vbCritical
    Err.Clear
End Sub

Public Sub PR_SMS_Close_CON()
    SMS_Connect
    If SMS.State = adStateOpen Then SMS.Close
End Sub

Public Sub PR_SMS_Open_CON()
    SMS_Connect
    If SMS.State = adStateclose Then SMS.Open
End Sub

Public Sub ADMS_Connect()
    On Error GoTo er_EH:
    Set con_ADMS = New ADODB.Connection
    con_ADMS.ConnectionTimeout = 0
    con_ADMS.Open ("Provider=SQLOLEDB.1;Data Source=10.227.241.27,1433;Network Library=DBMSSOCN;Initial Catalog=ProMIS_SX;User ID=ProMIS_User;Password=cosX@123;Trusted_Connection=False")
    Exit Sub
er_EH:
    MsgBox Err.Description, vbCritical
    Err.Clear
End Sub

