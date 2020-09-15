<%@ LANGUAGE=VBSCRIPT %>
<% 'On Error Resume Next %>
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<!-- <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8"> -->
</head>
<body>
<%
'ASP & Baspのセットアップ確認
Response.Write "<p>Hello Classic ASP!</p>"
Set objBASP21 = Server.CreateObject("basp21")
ver = objBASP21.Version()

Response.Write ver

%>
<br>
<%
'

set cn = Server.CreateObject ("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
cn.Open "dsn=PostgreSQL10;uid=postgres;pwd=secret"
'cn.CursorLocation = 3

set rs = cn.Execute("select * from mybook")
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
</body>
</html>