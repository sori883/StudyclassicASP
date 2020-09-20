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

  <!--  #include file="../inc/db.inc" -->

  <%
    sql = "select * from mybook ORDER BY id ASC"

    set cn = Server.CreateObject ("ADODB.Connection")
    Set rs = Server.CreateObject("ADODB.Recordset")
    cn.Open "dsn=PostgreSQL10;uid=postgres;pwd=secret"
    cn.CursorLocation = 3
  
    set rs = cn.Execute(sql)

    id = ""
    id = rs("id")
  %>
  <div>
    <h4>登録フォーム</h4>
    <form action="./create.asp" method="post" name="createForm">
      <label for="nameInput">名前</label>
      <input id="nameInput" type="text" name="username">
      <input type="submit" value="登録">
    </form>
  </div>

  <div>
    <h4>更新フォーム</h4>
    <form action="./update.asp" method="post" name="updateForm">
      <label for="nameInput">名前</label>
      <input id="nameInput" type="text" name="username">
      <input type="hidden" name="id" value="<%=id%>">
      <input type="submit" value="更新">
    </form>
  </div>

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

  <a href="delete.asp?id=<%=id%>"><%=id%>番目を削除する</a>


</body>
</html>