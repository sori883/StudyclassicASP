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
    Response.Cookies("param") = "�N�b�L�["
    Session("param") = "�Z�b�V����"
    Application("param") = "�A�v���P�[�V����"

  %>

  <a href="get.asp">get.asp�ɑJ��</a>
</body>
</html>