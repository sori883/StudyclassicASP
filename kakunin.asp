<%@ LANGUAGE=VBSCRIPT %>
<% 'On Error Resume Next %>
<!DOCTYPE html>
<html>
<head>
<meta charset="shift_jis">
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<!-- <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8"> -->
</head>
<body>
<%
' ASP & Basp�̃Z�b�g�A�b�v�m�F
Response.Write "<p>Hello Classic ASP!</p>"
Set objBASP21 = Server.CreateObject("basp21")
ver = objBASP21.Version()

Response.Write ver

%>
<br>
<%
'DB�ڑ��m�F

set cn = Server.CreateObject ("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
cn.Open "dsn=PostgreSQL10;uid=postgres;pwd=secret"
' count���g�p����ꍇ�ɕK�v
cn.CursorLocation = 3

' SQL���s
set rs = cn.Execute("select * from mybook")

Response.write "���U���g�̐�" & CStr(rs.RecordCount)

Response.write "<br>"

response.write("-------------<br>")
  do until rs.EOF
    response.write(rs("id") & "<br>")
    response.write(rs("name") & "<br>")
    rs.MoveNext
  loop
response.write("-------------<br>")

rs.Close
cn.Close



%>
<h4>���ϐ�</h4>
<%
For Each name In Request.ServerVariables
  Response.Write name & " = " & Request.ServerVariables(name) & "<br>"
Next
%>

</body>
</html>