<%
Class DbClass
  Private cn

  ' コンストラクタ
  Private Sub Class_Initialize()
    set cn = Server.CreateObject ("ADODB.Connection")
    cn.Open "dsn=PostgreSQL10;uid=postgres;pwd=secret"
    cn.CursorLocation = 3
  End Sub

  ' DB検索用
  Public Function DbSelecter(sql)
    Set DbSelecter = Server.CreateObject("ADODB.Recordset")
    set DbSelecter = cn.Execute(sql)
  End Function

  Public Function DbUpdater(sql)
    cn.execute(sql)
  End Function

  ' デストラクタ
  Private Sub Class_Terminate()
    cn.Close
  End Sub
End Class
%>