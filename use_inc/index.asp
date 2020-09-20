<% @LANGUAGE=VBSCRIPT %>
<% 'On Error Resume Next %>
<!DOCTYPE html>
<html>
<head>
<meta charset="shift_jis">
<meta http-equiv="X-UA-Compatible" content="IE=edge">
</head>
<body>
  <%
      error = session("error_create")
      If Not error = "" Then
        response.write error
        session.contents.remove("error_create")
      End If
  %>

  <!--  #include file="../inc/db_function.inc" -->

  <%
    sql = "select * from mybook ORDER BY id ASC"
    set rs = dbSelecter(sql)

    id = ""
    id = rs("id")
  %>
  <div>
  <%
    Response.write "<br>"
    Response.write "リザルト数" & CStr(rs.RecordCount)
    Response.write "<br>"

    response.write("-------------<br>")
    do until rs.EOF
      response.write(rs("id") & ",")
      response.write(rs("name") & "<br>")
      rs.MoveNext
    loop
    response.write("-------------<br>")
  %>



</body>
</html>