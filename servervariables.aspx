
<h3><% Response.Write (System.Environment.MachineName) %>'s Server Variables</h3>

<%
For Each var as String in Request.ServerVariables
  Response.Write(var & " " & Request(var) & "<br>")
Next
%>

<% Response.WriteFile ("footer.inc") %>