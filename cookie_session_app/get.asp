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
    c = Request.Cookies("param")
    s = Session("param")
    a = Application("param")

    response.write c & "<br>"
    response.write s & "<br>"
    response.write a & "<br>"
  %>
</body>
</html>