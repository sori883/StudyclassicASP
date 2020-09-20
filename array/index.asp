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
  配列は以下の通り<br>
  hoge = Array("どらえもん","のびた","ポプテピピック")
</p>

<%
  hoge = Array("どらえもん","のびた","ポプテピピック")
  response.write "<p>全件表示</p>"
  For i =0 To Ubound(hoge)
    response.write hoge(i) & "<br>"
  Next

  '配列の大きさを変更(既存の値は残したまま)
  Redim Preserve hoge(20)

  For i =0 To Ubound(hoge)
    response.write i
    response.write hoge(i) & "<br>"
  Next

  '文字列から配列を生成
  strMoji = "i am iron man"
  mojis = Split(strMoji)

  For i =0 To Ubound(mojis)
    response.write i
    response.write mojis(i) & "<br>"
  Next
%>
</body>
</html>