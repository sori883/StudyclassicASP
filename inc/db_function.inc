<%
Function DbSelecter(sql)
  set cn = Server.CreateObject ("ADODB.Connection")
  Set DbSelecter = Server.CreateObject("ADODB.Recordset")
  cn.Open "dsn=PostgreSQL10;uid=postgres;pwd=secret"
  cn.CursorLocation = 3

  set DbSelecter = cn.Execute(sql)
End Function
%>