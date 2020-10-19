<% @LANGUAGE=VBSCRIPT %>
<%
Set obj = Server.CreateObject("Excel.Application.16")
obj.Workbooks.Open("C:\inetpub\wwwroot\file\book.xls")

Set objWorksheet = obj.Worksheets("Sheet1")
response.write objWorksheet.Cells(1, 2).Value
%>