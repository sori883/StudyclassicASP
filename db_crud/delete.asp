<% @ LANGUAGE=VBSCRIPT %>
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
    id = Request.QueryString("id")
    
    set cn = Server.CreateObject ("ADODB.Connection")
    cn.Open "dsn=PostgreSQL10;uid=postgres;pwd=secret"

    sql = "DELETE FROM mybook WHERE id = "& id

    cn.execute(sql)

    cn.Close

    Response.Redirect "./index.asp"
  %>
</body>
</html>