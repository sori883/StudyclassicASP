set cn = Server.CreateObject ("ADODB.Connection")
cn.Open "dsn=PostgreSQL30;uid=postgres;pwd=secret"
cn.CursorLocation = 3

cn.Close