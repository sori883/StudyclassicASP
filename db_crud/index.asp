<% @ LANGUAGE=VBSCRIPT %>
<% 'On Error Resume Next %>
<!DOCTYPE html>
<html>
<head>
<meta charset="shift_jis">
<meta http-equiv="X-UA-Compatible" content="IE=edge">
</head>
<body>
  <%
    id = ""

    set cn = Server.CreateObject ("ADODB.Connection")
    Set rs = Server.CreateObject("ADODB.Recordset")
    cn.Open "dsn=PostgreSQL10;uid=postgres;pwd=secret"
    cn.CursorLocation = 3

    set rs = cn.Execute("select * from mybook ORDER BY id ASC")
    ' ��ԍŏ���ID���Z�b�g
    id = rs("id")
  %>
  <div>
    <h4>�o�^�p�t�H�[��</h4>
    <form action="./create.asp" method="post" name="createForm">
      <label for="nameInput">���O</label>
      <input id="nameInput" type="text" name="username">
      <input type="submit" value="�o�^">
    </form>
  </div>

  <div>
    <h4>�X�V�p�t�H�[��</h4>
    <form action="./update.asp" method="post" name="updateForm">
      <label for="nameInput">���O</label>
      <input id="nameInput" type="text" name="username">
      <input type="hidden" name="id" value="<%=id%>">
      <input type="submit" value="�o�^">
    </form>
  </div>

  <%
    Response.write "<br>"
    Response.write "���U���g�̐�" & CStr(rs.RecordCount)
    Response.write "<br>"

    response.write("-------------<br>")
    do until rs.EOF
      response.write(rs("id") & ",")
      response.write(rs("name") & "<br>")
      rs.MoveNext
    loop
    response.write("-------------<br>")
  %>

  <a href="delete.asp?id=<%=id%>"><%=id%>�Ԗڂ��폜</a>


</body>
</html>