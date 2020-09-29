<%@ LANGUAGE="VBScript" %>
<%
    Response.ContentType = "application/json"
    Dim ret
    If(Request.QueryString("name") = "Hoge") then
        ret = "{""name"": ""Huga""}"
    End If
    Response.Write(ret)
%>