<% @LANGUAGE=VBSCRIPT %>
<% 'On Error Resume Next %>
<!DOCTYPE html>
<html>
<head>
<meta charset="shift_jis">
<meta http-equiv="X-UA-Compatible" content="IE=edge">
</head>
<body>
 <!--  #include file="../inc/db_class.inc" -->
 <%
    error = session("error_db")
    If Not error = "" Then
      response.write error
      session.contents.remove("error_db")
    End If
  %>
  <%
   Set dbObj = New DbClass

   sql = "select * from mybook ORDER BY id ASC"
   set rs = dbObj.DbSelecter(sql)
   id = rs("id")
  %>

  <%
  response.write("-------------<br>")
  do until rs.EOF
    response.write(rs("id") & ",")
    response.write(rs("name") & "<br>")
    rs.MoveNext
  loop
  response.write("-------------<br>")
%>

<h1>更新関係</h1>
<h4>登録フォーム</h4>
<form action="./update.asp?doing=create" method="post" name="createForm">
  <label for="nameInput">名前</label>
  <input id="nameInput" type="text" name="username">
  <input type="submit" value="更新">
</form>
</div>

<div>
<h4>更新フォーム</h4>
<form action="./update.asp?doing=update" method="post" name="updateForm">
  <label for="nameInput">名前</label>
  <input id="nameInput" type="text" name="username">
  <input type="hidden" name="id" value="<%=id%>">
  <input type="submit" value="更新">
</form>
</div>

<a href="update.asp?doing=delete&id=<%=id%>"><%=id%>番目を削除する</a>

</body>
</html>