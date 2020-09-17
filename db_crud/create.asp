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
    validate = True
    error = ""

    username = Request.Form("username")

    'バリデーション

    If username = "" Then
      validate = False
      error="空文字は登録出来ません"
    End If

    If validate Then
      set cn = Server.CreateObject ("ADODB.Connection")
      cn.Open "dsn=PostgreSQL10;uid=postgres;pwd=secret"

      sql = "INSERT INTO mybook (name) VALUES ('"& username &"') "

      cn.execute(sql)
      
      cn.Close
    Else
      session("error_create") = error
    End If

    Response.Redirect "./index.asp"
  %>
</body>
</html>