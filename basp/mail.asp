<%@ LANGUAGE=VBSCRIPT %>
<% 'On Error Resume Next %>
<%
  Set basp = Server.CreateObject("basp21")

  svname = "localhost:1025"
  mailto = "受信者<testto@exsample.com>"
  mailfrom = "送信者<testfrom@exsample.com>"
  subj = "件名"
  body = "本文" & vbCrLf  & "以上"

  rc = basp.SendMail(svname,mailto,mailfrom, subj,body, "")

  if rc <> "" then	' エラーチェック
  response.write(rc)
end if		

%>