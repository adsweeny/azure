<%@LANGUAGE="VBSCRIPT"%>
<%
Dim Connection
Dim Recordset
Dim SQL
Dim Server
Dim field

'declare the SQL statement that will query the database
SQL = "SELECT * FROM dbo.Users"

'create an instance of the ADO connection and recordset objects
Set Connection = CreateObject("ADODB.Connection")
Set Recordset = CreateObject("ADODB.Recordset")
%>

<% Response.WriteFile ("footer.inc") %>