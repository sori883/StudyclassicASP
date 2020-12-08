<% @LANGUAGE=VBSCRIPT %>
<%
set obj =  Server.CreateObject("WScript.shell")
strCmd = "java -cp ""C:\temp\Hello.jar""; sample/Hello > C:\temp\hoge.txt 2>&1"
Response.Write strCmd
hoge = obj.run(strCmd, 1, True)

Response.Write "hogehogehoge"
%>