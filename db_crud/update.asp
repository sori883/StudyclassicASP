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
    username = Request.Form("username")
    id = Request.Form("id")
    
    set cn = Server.CreateObject ("ADODB.Connection")
    cn.Open "dsn=PostgreSQL10;uid=postgres;pwd=secret"

    sql = "UPDATE mybook SET name = '" & username & "' WHERE id = " & id

    cn.execute(sql)

    cn.Close

    Response.Redirect "./index.asp"
  %>
</body>
</html>