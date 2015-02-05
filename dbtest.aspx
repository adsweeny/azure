<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/BEnroll.asp" -->
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_BEnroll_STRING
Recordset1_cmd.CommandText = "SELECT * FROM dbo.Contacts" 
Recordset1_cmd.Prepared = true

Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 10
Repeat1__index = 0
Recordset1_numRows = Recordset1_numRows + Repeat1__numRows
%>



<table border="1">
  <tr>
    <td>ContactID</td>
    <td>Contactfirstname</td>
    <td>contactlastname</td>
    <td>contactdept</td>
  </tr>
  <% While ((Repeat1__numRows <> 0) AND (NOT Recordset1.EOF)) %>
    <tr>
      <td><%=(Recordset1.Fields.Item("ContactID").Value)%></td>
      <td><%=(Recordset1.Fields.Item("Contactfirstname").Value)%></td>
      <td><%=(Recordset1.Fields.Item("contactlastname").Value)%></td>
      <td><%=(Recordset1.Fields.Item("contactdept").Value)%></td>
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
