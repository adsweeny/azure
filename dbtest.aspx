<%@LANGUAGE="VBSCRIPT"%>
<%
Const adOpenStatic = 3 
Const adLockOptimistic = 3 
 
Set objConnection = CreateObject("ADODB.Connection") 
Set objRecordSet = CreateObject("ADODB.Recordset") 
 
objConnection.Open _ 
    "Provider=SQLOLEDB;Data Source=tgcdbbhzjb.database.windows.net;" & _ 
        "Trusted_Connection=Yes;Initial Catalog=adsweeny;" & _ 
             "User ID=adsweeny;Password=q988crunKJC4fvqcvX18;" 
 
objRecordSet.Open "SELECT * FROM Users", _ 
        objConnection, adOpenStatic, adLockOptimistic 
 
objRecordSet.MoveFirst 
 
Wscript.Echo objRecordSet.RecordCount 
%>

<% Response.WriteFile ("footer.inc") %>