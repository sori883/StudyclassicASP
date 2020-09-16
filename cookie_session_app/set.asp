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
    Response.Cookies("param") = "クッキー"
    Session("param") = "セッション"
    Application("param") = "アプリケーション"

  %>

  <a href="get.asp">get.aspに遷移</a>
</body>
</html>