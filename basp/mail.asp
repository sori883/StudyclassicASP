<%@ LANGUAGE=VBSCRIPT %>
<% 'On Error Resume Next %>
<%
  Set basp = Server.CreateObject("basp21")

  svname = "localhost:1025"
  mailto = "��M��<testto@exsample.com>"
  mailfrom = "���M��<testfrom@exsample.com>"
  subj = "����"
  body = "�{��" & vbCrLf  & "�ȏ�"

  rc = basp.SendMail(svname,mailto,mailfrom, subj,body, "")

  if rc <> "" then	' �G���[�`�F�b�N
  response.write(rc)
end if		

%>