<%@LANGUAGE="VBSCRIPT"%>
<%
option explicit

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

'open the connection to the database
Connection.Open "DSN=tgcdbbhzjb.database.windows.net;UID=adsweeny@tgcdbbhzjb;PWD=q988crunKJC4fvqcvX18;Database=adsweeny"

'Open the recordset object executing the SQL statement and return records 
Recordset.Open SQL,Connection

'first of all determine whether there are any records 
If Recordset.EOF Then 
wscript.echo "There are no records to retrieve; Check that you have the correct job number."
Else 
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 10
Repeat1__index = 0
Recordset1_numRows = Recordset1_numRows + Repeat1__numRows
%>



<table border="1">
  <tr>
    <td>UserID</td>
    <td>Userfirstname</td>
    <td>Userlastname</td>
    <td>username</td>
  </tr>
  <% While ((Repeat1__numRows <> 0) AND (NOT Recordset1.EOF)) %>
    <tr>
      <td><%=(Recordset1.Fields.Item("UserID").Value)%></td>
      <td><%=(Recordset1.Fields.Item("first").Value)%></td>
      <td><%=(Recordset1.Fields.Item("last").Value)%></td>
      <td><%=(Recordset1.Fields.Item("username").Value)%></td>
    </tr>
    <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Recordset1.MoveNext()
End While
%>
</table>
<%
Recordset1.Close()
Recordset1 = Nothing
%>

<% Response.WriteFile ("footer.inc") %>