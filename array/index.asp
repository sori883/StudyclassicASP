<% @LANGUAGE=VBSCRIPT %>
<% 'On Error Resume Next %>
<!DOCTYPE html>
<html>
<head>
<meta charset="shift_jis">
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<!-- <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8"> -->
</head>
<body>
<p>
  �z��͈ȉ��̒ʂ�<br>
  hoge = Array("�ǂ炦����","�̂т�","�|�v�e�s�s�b�N")
</p>

<%
  hoge = Array("�ǂ炦����","�̂т�","�|�v�e�s�s�b�N")
  response.write "<p>�S���\��</p>"
  For i =0 To Ubound(hoge)
    response.write hoge(i) & "<br>"
  Next

  '�z��̑傫����ύX(�����̒l�͎c�����܂�)
  Redim Preserve hoge(20)

  For i =0 To Ubound(hoge)
    response.write i
    response.write hoge(i) & "<br>"
  Next

  '�����񂩂�z��𐶐�
  strMoji = "i am iron man"
  mojis = Split(strMoji)

  For i =0 To Ubound(mojis)
    response.write i
    response.write mojis(i) & "<br>"
  Next
%>
</body>
</html>