<% @LANGUAGE=VBSCRIPT %>
<% 'On Error Resume Next %>
<!--  #include file="../inc/db_class.inc" -->
  <%

    doing = Request.QueryString("doing")

    sql = ""
    validate = False
    error = ""

    ' form variable
    username = ""
    id = ""

    Select Case doing
    Case "create"
    username = Request.Form("username")

      If username = "" Then
        validate = True
        error = "空文字は登録出来ません"
      Else
        sql = "INSERT INTO mybook (name) VALUES ('"& username &"') "
      End If

    Case "update"

      username = Request.Form("username")
        id = Request.Form("id")
        If username = "" Then
        validate = True
        error = "空文字は登録出来ません"
      Else
        sql = "UPDATE mybook SET name = '" & username & "' WHERE id = " & id
      End If

    Case "delete"
      id = Request.QueryString("id")
      If id = "" Then 
        validate = True
        error = "IDが指定されていません"
      Else
        sql = "DELETE FROM mybook WHERE id = "& id
      End If
    Case Else
      validate = True
      error = "正常に処理が行われませんでした"
    End Select

    If validate Then
      session("error_db") = error
      Response.Redirect "./class.asp"
    End If


    Set dbObj = New DbClass
    dbObj.DbSelecter(sql)

    Response.Redirect "./class.asp"
  %>
</body>
</html>