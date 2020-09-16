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
' ASP & Basp
Response.Write "<p>Hello Classic ASP!</p>"
Set objBASP21 = Server.CreateObject("basp21")
ver = objBASP21.Version()

Response.Write ver

%>
<br>
<%
'DB
set cn = Server.CreateObject ("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
cn.Open "dsn=PostgreSQL10;uid=postgres;pwd=secret"
' countを使う場合は3にする
cn.CursorLocation = 3

' SQL
set rs = cn.Execute("select * from mybook")

Response.write "リザルト数" & CStr(rs.RecordCount)

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
<h4>環境変数</h4>
<%
For Each name In Request.ServerVariables
  Response.Write name & " = " & Request.ServerVariables(name) & "<br>"
Next
%>

</body>
</html>